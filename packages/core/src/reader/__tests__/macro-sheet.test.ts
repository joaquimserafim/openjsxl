import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { evaluateCell } from "../../formula/eval";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openXlsx, streamSheetRows } from "../workbook";
import {
	createWorksheetStructure,
	pushWorksheetClose,
	pushWorksheetOpen,
} from "../worksheet-structure";
import { isSheetRoot } from "../worksheet-structure-helpers";
import { rewriteParts } from "./shared-part-fixtures";

// F11.32 (D13): an Excel 4.0 macro sheet, whose root is `xm:macrosheet` for both the xlMacrosheet and
// the xlIntlMacrosheet relationship, reads and bridges exactly like the same content in a worksheet.
// The sheet gate records the root name it opened, so nested roots and the root's close follow it.
// Every other foreign root, such as a chartsheet that carries cells, stays rejected.

const FIXTURE = "macrosheet-xlm.xlsm";
const MACRO1 = "xl/macrosheets/sheet1.xml";
const MACRO2 = "xl/macrosheets/sheet2.xml";
const REFS = ["A1", "A2", "A3", "A4", "A5", "B1"];

// Every accessor's view of one sheet: rows, cells, styles, number formats, formulas and streamed rows.
async function view(bytes: Uint8Array, sheet: string) {
	const book = await openXlsx(bytes);
	const ws = book.sheet(sheet);
	const rows: string[][] = [];
	for await (const row of ws.rows())
		rows.push(row.cells.map((c) => `${c.ref}=${String(c.value)}`));
	const streamed: string[][] = [];
	for await (const row of streamSheetRows(bytes, sheet)) {
		streamed.push(row.cells.map((c) => `${c.ref}=${String(c.value)}`));
	}
	return {
		rows,
		streamed,
		cells: REFS.map((ref) => ws.cell(ref)),
		styles: REFS.map((ref) => ws.style(ref)),
		formats: REFS.map((ref) => ws.numberFormat(ref)),
		formulas: REFS.map((ref) => ws.formula(ref)),
	};
}

// The same parts with each macro-sheet root renamed to a worksheet root.
const asWorksheets = (bytes: Uint8Array, extra: (xml: string) => string = (xml) => xml) =>
	rewriteParts(bytes, {
		[MACRO1]: (xml) => extra(xml.replace(/xm:macrosheet/g, "worksheet")),
		[MACRO2]: (xml) => xml.replace(/xm:macrosheet/g, "worksheet"),
	});

describe("an Excel 4.0 macro sheet", () => {
	it("reads its cells, styles, number formats and formulas through every accessor", async () => {
		const bytes = await loadFixture(FIXTURE);
		const book = await openXlsx(bytes);
		expect(book.sheets.map((s) => s.name)).toEqual(["Macro1", "Macro2", "Chart1"]);
		const macro1 = await view(bytes, "Macro1");
		expect(macro1.rows).toEqual([["A1=42"], ["A2=TRUE"], ["A3=true"], ["A4=note"], ["A5=42"]]);
		expect(macro1.streamed).toEqual(macro1.rows);
		expect(macro1.cells.map((c) => [c.type, c.value])).toEqual([
			["number", 42],
			["string", "TRUE"],
			["boolean", true],
			["string", "note"],
			["number", 42],
			["empty", null],
		]);
		expect(macro1.formats.slice(0, 2)).toEqual(["0.000", "General"]);
		expect(macro1.styles[2]?.font?.bold).toBe(true);
		expect(macro1.formulas[4]).toBe("RETURN(A1)");
	});

	it("reads the same under the xlIntlMacrosheet relationship", async () => {
		const macro2 = await view(await loadFixture(FIXTURE), "Macro2");
		expect(macro2.rows).toEqual([["A1=7.5", "B1=intl"], ["A2=15"]]);
		expect(macro2.streamed).toEqual(macro2.rows);
		expect(macro2.formats[0]).toBe("0.000");
		expect(macro2.formulas[1]).toBe("RETURN(A1*2)");
	});

	it("reads exactly like the same content in a worksheet", async () => {
		const bytes = await loadFixture(FIXTURE);
		const twin = await asWorksheets(bytes);
		for (const sheet of ["Macro1", "Macro2"]) {
			expect(await view(bytes, sheet)).toEqual(await view(twin, sheet));
		}
	});

	it("is seen by the formula evaluator", async () => {
		const book = await openXlsx(await loadFixture(FIXTURE));
		expect(await evaluateCell(book, "Macro1", "A1")).toBe(42);
		expect(await evaluateCell(book, "Macro2", "B1")).toBe("intl");
	});

	it("bridges to a regular worksheet with its cells, styles and formulas", async () => {
		const book = await openXlsx(await loadFixture(FIXTURE));
		const rewritten = await writeXlsx(await workbookToInput(book));
		expect((await view(rewritten, "Macro1")).rows).toEqual([
			["A1=42"],
			["A2=TRUE"],
			["A3=true"],
			["A4=note"],
			["A5=42"],
		]);
		expect((await view(rewritten, "Macro2")).rows).toEqual([["A1=7.5", "B1=intl"], ["A2=15"]]);
		for (const sheet of ["Macro1", "Macro2"]) {
			const before = await view(await loadFixture(FIXTURE), sheet);
			const after = await view(rewritten, sheet);
			expect(after.rows).toEqual(before.rows);
			expect(after.formats).toEqual(before.formats);
			expect(after.formulas).toEqual(before.formulas);
			expect(after.styles.map((s) => s?.font?.bold ?? false)).toEqual(
				before.styles.map((s) => s?.font?.bold ?? false),
			);
		}
	});

	it("stops at its own closing tag: a row after it is not read", async () => {
		const trailing = (xml: string) =>
			xml.replace(
				"</sheetData></xm:macrosheet>",
				'</xm:macrosheet><row r="9"><c r="A9"><v>777</v></c></row>',
			);
		const bytes = await rewriteParts(await loadFixture(FIXTURE), { [MACRO1]: trailing });
		const macro = await view(bytes, "Macro1");
		expect(macro.rows).toEqual([["A1=42"], ["A2=TRUE"], ["A3=true"], ["A4=note"], ["A5=42"]]);
		expect((await openXlsx(bytes)).sheet("Macro1").cell("A9").value).toBeNull();
		expect(macro).toEqual(
			await view(
				await asWorksheets(await loadFixture(FIXTURE), (xml) =>
					xml.replace(
						"</sheetData></worksheet>",
						'</worksheet><row r="9"><c r="A9"><v>777</v></c></row>',
					),
				),
				"Macro1",
			),
		);
	});

	it("still reads its cells past a nested macrosheet, whose own cells stay unread", async () => {
		const nested = (root: string) => (xml: string) =>
			xml.replace(
				"<dimension",
				`<sheetPr><extLst><ext uri="{X}"><${root}><sheetData><row r="9"><c r="A9"><v>666</v></c></row></sheetData></${root}></ext></extLst></sheetPr><dimension`,
			);
		const bytes = await rewriteParts(await loadFixture(FIXTURE), {
			[MACRO1]: nested("xm:macrosheet"),
		});
		const macro = await view(bytes, "Macro1");
		expect(macro.rows).toEqual([["A1=42"], ["A2=TRUE"], ["A3=true"], ["A4=note"], ["A5=42"]]);
		expect(macro).toEqual(
			await view(
				await asWorksheets(await loadFixture(FIXTURE), nested("worksheet")),
				"Macro1",
			),
		);
	});
});

describe("other sheet roots", () => {
	it("leave a chartsheet that carries cells without cells, as before", async () => {
		const chart = await view(await loadFixture(FIXTURE), "Chart1");
		expect(chart.rows).toEqual([]);
		expect(chart.streamed).toEqual([]);
		expect(chart.cells.every((c) => c.type === "empty")).toBe(true);
	});
});

describe("isSheetRoot", () => {
	it("admits the worksheet and macro-sheet roots only", () => {
		expect(isSheetRoot("worksheet")).toBe(true);
		expect(isSheetRoot("macrosheet")).toBe(true);
		for (const name of [
			"chartsheet",
			"dialogsheet",
			"wrapper",
			"Worksheet",
			"xm:macrosheet",
			"",
		]) {
			expect(isSheetRoot(name)).toBe(false);
		}
	});
});

describe("the sheet gate's root", () => {
	it("records the root it opened, prefixed or not, and closes on that root only", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "xm:macrosheet", false);
		expect([state.phase, state.rootName]).toEqual(["worksheet", "macrosheet"]);
		// A close of the other sheet-root name is an ordinary close.
		expect(pushWorksheetClose(state, "worksheet")).toBeUndefined();
		expect(state.phase).toBe("worksheet");
		pushWorksheetClose(state, "xm:macrosheet");
		expect(state.phase).toBe("closed");
	});

	it("counts nested roots of the opened name only", () => {
		const macro = createWorksheetStructure();
		pushWorksheetOpen(macro, "xm:macrosheet", false);
		pushWorksheetOpen(macro, "sheetPr", false);
		pushWorksheetOpen(macro, "xm:macrosheet", false);
		pushWorksheetOpen(macro, "worksheet", false);
		expect(macro.rootSameNameDepth).toBe(1);
		pushWorksheetClose(macro, "worksheet");
		pushWorksheetClose(macro, "xm:macrosheet");
		pushWorksheetClose(macro, "sheetPr");
		expect([macro.phase, macro.rootSameNameDepth]).toEqual(["worksheet", 0]);

		// A worksheet with a nested macrosheet keeps reading past it.
		const sheet = createWorksheetStructure();
		pushWorksheetOpen(sheet, "worksheet", false);
		pushWorksheetOpen(sheet, "sheetPr", false);
		pushWorksheetOpen(sheet, "xm:macrosheet", false);
		pushWorksheetClose(sheet, "xm:macrosheet");
		pushWorksheetClose(sheet, "sheetPr");
		pushWorksheetOpen(sheet, "sheetData", false);
		expect(pushWorksheetOpen(sheet, "row", false)).toBe("row");
	});

	it("closes on a self-closing macro-sheet root, and on any other root", () => {
		const roots: readonly (readonly [string, boolean])[] = [
			["xm:macrosheet", true],
			["chartsheet", false],
			["dialogsheet", false],
		];
		for (const [name, selfClosing] of roots) {
			const state = createWorksheetStructure();
			pushWorksheetOpen(state, name, selfClosing);
			expect([state.phase, state.rootName]).toEqual(["closed", undefined]);
		}
	});
});
