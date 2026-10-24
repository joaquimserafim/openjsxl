import { listFixtures, loadFixture } from "@openjsxl/fixtures";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { openZip } from "../../zip";
import { detectSpreadsheetFormat } from "../detect";
import { openXlsx, streamSheetRows } from "../workbook";
import { openXlsb } from "../xlsb";
import {
	declaredSheets,
	rewriteParts,
	TABLE_REL,
	threeSheetWorkbook,
} from "./shared-part-fixtures";

// F11.95: one open reads each package part at most once. A later sheet whose worksheet part is
// already loaded is dropped; a comments or table part already read is skipped. Every read goes
// through a counting wrapper around openZip, so the pins count reads per part directly: a doubling
// ratio cannot tell N reads of one part from one.

const reads = vi.hoisted(() => new Map<string, number>());

vi.mock("../../zip", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../zip")>();
	return {
		...actual,
		openZip: (...args: Parameters<typeof actual.openZip>) => {
			const archive = actual.openZip(...args);
			const count = (name: string) => reads.set(name, (reads.get(name) ?? 0) + 1);
			return {
				...archive,
				read: (name: string) => {
					count(name);
					return archive.read(name);
				},
				readStream: (name: string) => {
					count(name);
					return archive.readStream(name);
				},
			};
		},
	};
});

const base = await threeSheetWorkbook();
const sharedSheet = await rewriteParts(base, {
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="worksheets/sheet3.xml"', 'Target="worksheets/sheet1.xml"'),
});
const sheetOnWorkbook = await rewriteParts(base, {
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="worksheets/sheet3.xml"', 'Target="workbook.xml"'),
});
const sharedTable = await rewriteParts(base, {
	"xl/worksheets/_rels/sheet1.xml.rels": (t) =>
		t.replace(
			"</Relationships>",
			`<Relationship Id="rId9" Type="${TABLE_REL}" Target="../tables/table1.xml"/></Relationships>`,
		),
});
const sharedComments = await rewriteParts(base, {
	"xl/worksheets/_rels/sheet2.xml.rels": (t) => t.replace("../comments2.xml", "../comments1.xml"),
});
const sharedXlsbSheet = await rewriteParts(await loadFixture("xlsb-basic.xlsb"), {
	"xl/_rels/workbook.bin.rels": (t) =>
		t.replace('Target="worksheets/sheet3.bin"', 'Target="worksheets/sheet2.bin"'),
});

// S2 is renamed to repeat S1's name, so it is dropped for its name and claims no part; S3 then
// names S2's part and must load.
const nameDropThenPart = await rewriteParts(base, {
	"xl/workbook.xml": (t) => t.replace('name="S2"', 'name="S1"'),
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="worksheets/sheet3.xml"', 'Target="worksheets/sheet2.xml"'),
});
// S2 and S3 are both named B; S2 names S1's part, so it is dropped for its part and reserves no name.
const partDropThenName = await rewriteParts(base, {
	"xl/workbook.xml": (t) => t.replace('name="S2"', 'name="B"').replace('name="S3"', 'name="B"'),
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="worksheets/sheet2.xml"', 'Target="worksheets/sheet1.xml"'),
});
// S1 names S2's rels part as a table, so S1 claims it first and S2 loads without its rels.
const relsTakenByTable = await rewriteParts(base, {
	"xl/worksheets/_rels/sheet1.xml.rels": (t) =>
		t.replace(
			"</Relationships>",
			`<Relationship Id="rId9" Type="${TABLE_REL}" Target="_rels/sheet2.xml.rels"/></Relationships>`,
		),
});
const basic = await loadFixture("basic.xlsx");
const workbookLevelTargets: [target: string, part: string][] = [
	["sharedStrings.xml", "xl/sharedStrings.xml"],
	["styles.xml", "xl/styles.xml"],
	["_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels"],
	["../_rels/.rels", "_rels/.rels"],
	["../[Content_Types].xml", "[Content_Types].xml"],
];
// basic.xlsx with its one sheet naming each workbook-level part in turn.
const sheetOnPart = new Map<string, Uint8Array>();
for (const [target] of workbookLevelTargets) {
	sheetOnPart.set(
		target,
		await rewriteParts(basic, {
			"xl/_rels/workbook.xml.rels": (t) =>
				t.replace('Target="worksheets/sheet1.xml"', `Target="${target}"`),
		}),
	);
}
const sheetOnTheme = await rewriteParts(await loadFixture("any_sheets.xlsx"), {
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="worksheets/sheet3.xml"', 'Target="theme/theme1.xml"'),
});
// The styles relationship names the shared-strings part, so styles read as absent.
const stylesOnStrings = await rewriteParts(basic, {
	"xl/_rels/workbook.xml.rels": (t) =>
		t.replace('Target="styles.xml"', 'Target="sharedStrings.xml"'),
});

beforeEach(() => reads.clear());

describe("a worksheet part named by two sheets", () => {
	it("loads the first sheet and drops the later one with its scoped names", async () => {
		const book = await openXlsx(sharedSheet);
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "S2"]);
		expect(book.sheet("S1").cell("A1").value).toBe("a");
		expect(() => book.sheet("S3")).toThrow(expect.objectContaining({ code: "no-such-sheet" }));
		expect(book.definedNames).toEqual([
			{ name: "Everywhere", refersTo: "S1!$A$1" },
			{ name: "OnFirst", refersTo: "S1!$A$1", localSheetId: 0 },
		]);
		expect(reads.get("xl/worksheets/sheet1.xml")).toBe(1);
	});

	it("streams only the loaded sheets", async () => {
		await expect(streamSheetRows(sharedSheet, "S3").next()).rejects.toMatchObject({
			code: "no-such-sheet",
		});
		const first = await streamSheetRows(sharedSheet, "S1").next();
		expect(first.done === false && first.value.cells[0]?.value).toBe("a");
	});

	it("drops a sheet whose part is a workbook-level part this open already read", async () => {
		expect(openZip(base).has("xl/workbook.xml")).toBe(true);
		const book = await openXlsx(sheetOnWorkbook);
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "S2"]);
		expect(reads.get("xl/workbook.xml")).toBe(1);
	});

	it("lets a later sheet load a part that a sheet dropped for its name never claimed", async () => {
		const book = await openXlsx(nameDropThenPart);
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "S3"]);
		expect(book.sheet("S3").cell("A1").value).toBe(3);
	});

	it("lets a later sheet take the name of a sheet dropped for its part", async () => {
		const book = await openXlsx(partDropThenName);
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "B"]);
		expect(book.sheet("B").cell("A1").value).toBe(4);
		expect(book.definedNames.map((n) => [n.name, n.localSheetId])).toEqual([
			["Everywhere", undefined],
			["OnFirst", 0],
			["OnThird", 1],
		]);
	});

	it.each(workbookLevelTargets)(
		"drops a sheet whose part is the workbook-level part %s",
		async (target, part) => {
			expect((await openXlsx(sheetOnPart.get(target) ?? basic)).sheets).toEqual([]);
			expect(reads.get(part)).toBe(1);
		},
	);

	it("drops a sheet whose part is the theme part", async () => {
		const book = await openXlsx(sheetOnTheme);
		expect(book.sheets.map((s) => s.name)).toEqual(["Visible", "Hidden", "Chart"]);
		expect(reads.get("xl/theme/theme1.xml")).toBe(1);
	});

	it("loads the first of two xlsb sheets on one part", async () => {
		const original = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheets.map(
			(s) => s.name,
		);
		reads.clear();
		const book = await openXlsb(sharedXlsbSheet);
		expect(original).toHaveLength(3);
		expect(book.sheets.map((s) => s.name)).toEqual(original.slice(0, 2));
		expect(reads.get("xl/worksheets/sheet2.bin")).toBe(1);
	});
});

describe("comments and table parts named twice", () => {
	it("keeps one table when two table relationships target one part", async () => {
		const book = await openXlsx(sharedTable);
		expect(book.sheet("S1").tables.map((t) => t.name)).toEqual(["Sales"]);
		expect(reads.get("xl/tables/table1.xml")).toBe(1);
	});

	it("leaves a sheet without its rels when another role read that part first", async () => {
		const book = await openXlsx(relsTakenByTable);
		expect(book.sheet("S2").comments).toEqual([]);
		expect(book.sheet("S1").comments.map((c) => c.text)).toEqual(["c1"]);
		expect(reads.get("xl/worksheets/_rels/sheet2.xml.rels")).toBe(1);
	});

	it("reads a styles relationship that names the shared-strings part as no styles", async () => {
		const book = await openXlsx(stylesOnStrings);
		expect(book.sheet("Sheet1").cell("A1").value).toBe("hello");
		expect(book.sheet("Sheet1").cell("C1").type).toBe("number");
		expect(reads.get("xl/sharedStrings.xml")).toBe(1);
	});

	it("gives a shared comments part to the first sheet only", async () => {
		const book = await openXlsx(sharedComments);
		expect(book.sheet("S1").comments.map((c) => c.text)).toEqual(["c1"]);
		expect(book.sheet("S2").comments).toEqual([]);
		expect(reads.get("xl/comments1.xml")).toBe(1);
	});
});

describe("read counts", () => {
	it("reads every part at most once per open, for the shared-part shapes and the corpus", async () => {
		const inputs: [string, Uint8Array][] = [
			["shared sheet", sharedSheet],
			["sheet on the workbook part", sheetOnWorkbook],
			["shared table", sharedTable],
			["shared comments", sharedComments],
			["shared xlsb sheet", sharedXlsbSheet],
			["name drop then part", nameDropThenPart],
			["part drop then name", partDropThenName],
			["rels taken by a table", relsTakenByTable],
			["sheet on the theme part", sheetOnTheme],
			["styles on the shared strings", stylesOnStrings],
		];
		for (const [target, bytes] of sheetOnPart) inputs.push([`sheet on ${target}`, bytes]);
		for (const name of listFixtures()) inputs.push([name, await loadFixture(name)]);
		let opened = 0;
		for (const [name, bytes] of inputs) {
			const format = await detectSpreadsheetFormat(bytes);
			if (format !== "xlsx" && format !== "xlsb") continue;
			reads.clear();
			const [result] = await Promise.allSettled([
				format === "xlsb" ? openXlsb(bytes) : openXlsx(bytes),
			]);
			if (result?.status === "fulfilled") opened++;
			const repeated = [...reads].filter(([, count]) => count > 1);
			expect(repeated, name).toEqual([]);
		}
		expect(opened).toBeGreaterThanOrEqual(40);
	});
});

describe("committed corpus", () => {
	it("loads every declared sheet with its tables and comments", async () => {
		let checked = 0;
		for (const name of listFixtures()) {
			const bytes = await loadFixture(name);
			if ((await detectSpreadsheetFormat(bytes)) !== "xlsx") continue;
			const [opened] = await Promise.allSettled([openXlsx(bytes)]);
			if (opened?.status !== "fulfilled") continue;
			const book = opened.value;
			const declared = await declaredSheets(bytes);
			// No corpus workbook shares a part, so nothing may be dropped or skipped.
			expect(new Set(declared.map((s) => s.path)).size, name).toBe(declared.length);
			expect(book.sheets.length, name).toBe(declared.length);
			book.sheets.forEach((info, i) => {
				const sheet = book.sheet(info.name);
				expect(sheet.tables.length, `${name} ${info.name}`).toBe(
					declared[i]?.tables.length,
				);
				expect(sheet.comments.length, `${name} ${info.name}`).toBe(
					declared[i]?.commentCount,
				);
			});
			checked++;
		}
		expect(checked).toBeGreaterThanOrEqual(35);
	});
});
