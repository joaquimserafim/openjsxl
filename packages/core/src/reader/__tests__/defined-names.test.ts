import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { writeZip } from "../../writer/zip";
import { openXlsx, streamSheetRows } from "../workbook";

// F10.1 — read openpyxl-authored defined names verbatim, and round-trip them through the bridge.
// The fixture (openpyxl 3.1.5) carries a global range, a global constant, a global HIDDEN name, a
// sheet-local name, and a built-in _xlnm.Print_Area — see packages/fixtures/data/README.md.

const enc = new TextEncoder();
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";

const EXPECTED = [
	{ name: "Amounts", refersTo: "Data!$B$1:$B$3" },
	{ name: "TaxRate", refersTo: "0.2" },
	{ name: "SecretName", refersTo: "Data!$A$1", hidden: true },
	{ name: "FirstItem", refersTo: "Data!$A$1", localSheetId: 0 },
	{ name: "_xlnm.Print_Area", refersTo: "'Data'!$A$1:$B$3", localSheetId: 0 },
];

describe("reader — openpyxl defined names", () => {
	it("reads every name verbatim, in document order (global, constant, hidden, sheet-local, built-in)", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-defined-names.xlsx"));
		expect(wb.definedNames).toEqual(EXPECTED);
	});

	it("round-trips the names losslessly through the bridge", async () => {
		const before = await openXlsx(await loadFixture("openpyxl-defined-names.xlsx"));
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		expect(after.definedNames).toEqual(EXPECTED);
	});
});

// A minimal workbook declaring THREE sheets (S1/S2/S3) where S2's worksheet part is OMITTED, so the
// reader drops S2 and the survivors [S1, S3] compact to loaded indices 0 and 1. Four defined names
// exercise the F10.1 localSheetId reconciliation: a global name, one scoped to a surviving sheet at
// its original index, one scoped to the DROPPED sheet, and one scoped to a sheet whose index SHIFTS.
function droppedSheetWorkbook(): Promise<Uint8Array> {
	const sheet = `<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData/></worksheet>`;
	const workbook =
		`<?xml version="1.0"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets>` +
		`<sheet name="S1" sheetId="1" r:id="rId1"/>` +
		`<sheet name="S2" sheetId="2" r:id="rId2"/>` +
		`<sheet name="S3" sheetId="3" r:id="rId3"/>` +
		`</sheets><definedNames>` +
		`<definedName name="GlobalName">S1!$A$1</definedName>` +
		`<definedName name="FirstScope" localSheetId="0">S1!$A$1</definedName>` + // survives at index 0
		`<definedName name="DroppedScope" localSheetId="1">S2!$A$1</definedName>` + // scope sheet dropped
		`<definedName name="ShiftedScope" localSheetId="2">S3!$A$1</definedName>` + // shifts 2 → 1
		`</definedNames></workbook>`;
	const wbRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/>` +
		`<Relationship Id="rId2" Type="${NS_R}/worksheet" Target="worksheets/sheet2.xml"/>` + // dangles (no part)
		`<Relationship Id="rId3" Type="${NS_R}/worksheet" Target="worksheets/sheet3.xml"/>` +
		`</Relationships>`;
	const rootRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	const types =
		`<?xml version="1.0"?><Types xmlns="${CT}">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
		`</Types>`;
	return writeZip([
		{ name: "[Content_Types].xml", data: enc.encode(types) },
		{ name: "_rels/.rels", data: enc.encode(rootRels) },
		{ name: "xl/workbook.xml", data: enc.encode(workbook) },
		{ name: "xl/_rels/workbook.xml.rels", data: enc.encode(wbRels) },
		{ name: "xl/worksheets/sheet1.xml", data: enc.encode(sheet) },
		// xl/worksheets/sheet2.xml is intentionally ABSENT → S2 is dropped on load.
		{ name: "xl/worksheets/sheet3.xml", data: enc.encode(sheet) },
	]);
}

// Five raw sheet positions exercise the F11.4 collision rule after ST_Xstring decoding:
// A is readable first; raw `_x0041_` and `_x0061_` collapse to A/a and are dropped; the protected
// spelling remains the distinct literal `_x0041_`; Tail survives after both drops. Local defined
// names prove that collision drops use the same raw→loaded remap as a dangling worksheet part.
function encodedSheetCollisionWorkbook(
	omittedSheetIndexes: ReadonlySet<number> = new Set(),
	relationshipOverrides: ReadonlyMap<number, "missing" | "external"> = new Map(),
): Promise<Uint8Array> {
	const rawNames = ["A", "_x0041_", "_x0061_", "_x005F_x0041_", "Tail"];
	const worksheet = (value: number) =>
		`<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData>` +
		`<row r="1"><c r="A1"><v>${value}</v></c></row></sheetData></worksheet>`;
	const workbook =
		`<?xml version="1.0"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets>` +
		rawNames
			.map(
				(name, index) =>
					`<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
			)
			.join("") +
		`</sheets><definedNames>` +
		`<definedName name="GlobalName">42</definedName>` +
		`<definedName name="FirstScope" localSheetId="0">A!A1</definedName>` +
		`<definedName name="DroppedExactScope" localSheetId="1">A!A1</definedName>` +
		`<definedName name="DroppedCaseScope" localSheetId="2">a!A1</definedName>` +
		`<definedName name="LiteralScope" localSheetId="3">'_x005F_x0041_'!A1</definedName>` +
		`<definedName name="TailScope" localSheetId="4">Tail!A1</definedName>` +
		`</definedNames></workbook>`;
	const wbRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		rawNames
			.flatMap((_name, index) => {
				const override = relationshipOverrides.get(index);
				if (override === "missing") return [];
				const target =
					override === "external"
						? "https://example.invalid/sheet.xml"
						: `worksheets/sheet${index + 1}.xml`;
				const targetMode = override === "external" ? ' TargetMode="External"' : "";
				return [
					`<Relationship Id="rId${index + 1}" Type="${NS_R}/worksheet" Target="${target}"${targetMode}/>`,
				];
			})
			.join("") +
		`</Relationships>`;
	const rootRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	const types =
		`<?xml version="1.0"?><Types xmlns="${CT}">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
		`</Types>`;
	return writeZip([
		{ name: "[Content_Types].xml", data: enc.encode(types) },
		{ name: "_rels/.rels", data: enc.encode(rootRels) },
		{ name: "xl/workbook.xml", data: enc.encode(workbook) },
		{ name: "xl/_rels/workbook.xml.rels", data: enc.encode(wbRels) },
		...rawNames.flatMap((_name, index) =>
			omittedSheetIndexes.has(index)
				? []
				: [
						{
							name: `xl/worksheets/sheet${index + 1}.xml`,
							data: enc.encode(worksheet(index + 1)),
						},
					],
		),
	]);
}

describe("reader — sheet-scoped defined names when a sheet is dropped (F10.1 reconciliation)", () => {
	it("remaps a shifted localSheetId, drops a name scoped to a dropped sheet, keeps globals", async () => {
		const wb = await openXlsx(await droppedSheetWorkbook());
		expect(wb.sheets.map((s) => s.name)).toEqual(["S1", "S3"]); // S2 dropped
		expect(wb.definedNames).toEqual([
			{ name: "GlobalName", refersTo: "S1!$A$1" },
			{ name: "FirstScope", refersTo: "S1!$A$1", localSheetId: 0 }, // unchanged
			{ name: "ShiftedScope", refersTo: "S3!$A$1", localSheetId: 1 }, // remapped 2 → 1 (DroppedScope gone)
		]);
	});

	it("re-writes without aborting — every surviving localSheetId is a valid index into the loaded sheets", async () => {
		// The bug this pins: before reconciliation the reader kept a localSheetId the writer then
		// rejected (0..0), aborting the whole rewrite. Now the round-trip completes.
		const wb = await openXlsx(await droppedSheetWorkbook());
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(wb)));
		expect(rewritten.definedNames).toEqual([
			{ name: "GlobalName", refersTo: "S1!$A$1" },
			{ name: "FirstScope", refersTo: "S1!$A$1", localSheetId: 0 },
			{ name: "ShiftedScope", refersTo: "S3!$A$1", localSheetId: 1 },
		]);
	});
});

describe("reader — decoded ST_Xstring sheet-name collisions (F11.4)", () => {
	it("keeps the first readable logical name and remaps scoped defined names around dropped aliases", async () => {
		const bytes = await encodedSheetCollisionWorkbook();
		const wb = await openXlsx(bytes);
		expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["A", "_x0041_", "Tail"]);
		expect(wb.sheet("A").cell("A1").value).toBe(1);
		expect(wb.sheet("_x0041_").cell("A1").value).toBe(4);
		expect(wb.sheet("Tail").cell("A1").value).toBe(5);
		expect(() => wb.sheet("a")).toThrow(/no sheet named/);
		expect(wb.definedNames).toEqual([
			{ name: "GlobalName", refersTo: "42" },
			{ name: "FirstScope", refersTo: "A!A1", localSheetId: 0 },
			{ name: "LiteralScope", refersTo: "'_x0041_'!A1", localSheetId: 1 },
			{ name: "TailScope", refersTo: "Tail!A1", localSheetId: 2 },
		]);

		// The streaming API shares loadWorkbook: logical selection must reach the same survivors.
		const firstValue = async (name?: string): Promise<unknown> => {
			for await (const row of streamSheetRows(bytes, name)) return row.cells[0]?.value;
			return undefined;
		};
		expect(await firstValue()).toBe(1);
		expect(await firstValue("A")).toBe(1);
		expect(await firstValue("_x0041_")).toBe(4);
		expect(await firstValue("Tail")).toBe(5);
		await expect(firstValue("a")).rejects.toMatchObject({ code: "no-such-sheet" });
	});

	it("reserves a decoded name only after its worksheet part is readable", async () => {
		// Raw sheet 0 (`A`) dangles, so raw sheet 1 (`_x0041_` → A) becomes the first readable A.
		const bytes = await encodedSheetCollisionWorkbook(new Set([0]));
		const wb = await openXlsx(bytes);
		expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["A", "_x0041_", "Tail"]);
		expect(wb.sheet("A").cell("A1").value).toBe(2);
		expect(wb.definedNames).toEqual([
			{ name: "GlobalName", refersTo: "42" },
			{ name: "DroppedExactScope", refersTo: "A!A1", localSheetId: 0 },
			{ name: "LiteralScope", refersTo: "'_x0041_'!A1", localSheetId: 1 },
			{ name: "TailScope", refersTo: "Tail!A1", localSheetId: 2 },
		]);
		let streamedValue: unknown;
		for await (const row of streamSheetRows(bytes, "A")) {
			streamedValue = row.cells[0]?.value;
			break;
		}
		expect(streamedValue).toBe(2);
	});

	it.each(["missing", "external"] as const)(
		"does not reserve a decoded name for a %s worksheet relationship",
		async (relationshipMode) => {
			const bytes = await encodedSheetCollisionWorkbook(
				new Set(),
				new Map([[0, relationshipMode]]),
			);
			const wb = await openXlsx(bytes);
			expect(wb.sheets.map((sheet) => sheet.name)).toEqual(["A", "_x0041_", "Tail"]);
			expect(wb.sheet("A").cell("A1").value).toBe(2);
			expect(wb.definedNames).toContainEqual({
				name: "DroppedExactScope",
				refersTo: "A!A1",
				localSheetId: 0,
			});
		},
	);

	it("bridges only the writer-legal survivors without losing the reconciled scopes", async () => {
		const before = await openXlsx(await encodedSheetCollisionWorkbook());
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		expect(after.sheets.map((sheet) => sheet.name)).toEqual(["A", "_x0041_", "Tail"]);
		expect(after.definedNames).toEqual(before.definedNames);
		expect(after.sheet("A").cell("A1").value).toBe(1);
		expect(after.sheet("_x0041_").cell("A1").value).toBe(4);
		expect(after.sheet("Tail").cell("A1").value).toBe(5);
	});
});
