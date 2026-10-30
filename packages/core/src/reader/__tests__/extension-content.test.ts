import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openXlsx, streamSheetRows } from "../workbook";
import { parseColumnProps, parseFormulas, parseRowProperties, readRows } from "../worksheet";

// F11.26 and F11.38: extension content inside a cell, a row or sheetPr is foreign to the worksheet.
// Its `<v>`, `<is>`, `<t>`, `<f>`, `<row>` and `<col>` elements, prefixed or not, never supply a
// cell value, a formula, row properties or a column, in the buffered reader, the streaming reader
// or the bridge. The cell's own direct `<v>`, `<is>` and `<f>` still read as before.

const FIXTURE = "extension-content.xlsx";

async function view(bytes: Uint8Array, sheet: string) {
	const ws = (await openXlsx(bytes)).sheet(sheet);
	const cells: Record<string, unknown> = {};
	const formulas: Record<string, string> = {};
	for await (const row of ws.rows()) {
		for (const cell of row.cells) {
			// rows() and cell() must agree on every cell.
			expect(ws.cell(cell.ref).value).toEqual(cell.value);
			cells[cell.ref] = cell.value;
			const formula = ws.formula(cell.ref);
			if (formula !== undefined) formulas[cell.ref] = formula;
		}
	}
	const streamed: Record<string, unknown> = {};
	for await (const row of streamSheetRows(bytes, sheet)) {
		for (const cell of row.cells) streamed[cell.ref] = cell.value;
	}
	return { cells, streamed, formulas, rowProperties: [...ws.rowProperties], columns: ws.columns };
}

const bridged = async (bytes: Uint8Array) =>
	writeXlsx(await workbookToInput(await openXlsx(bytes)));

const VALUES = {
	A1: 7,
	A2: 7,
	A3: "good",
	A4: "good",
	A5: "good",
	A6: 7,
	A7: 7,
	A8: "good",
	A9: "漢字",
	A10: "shared",
	A11: "",
};
const FORMULAS = { A2: "1+1", A3: "2+2", A4: "SUM(A1:A3)", A5: "A1*2", A6: "A2*2" };

describe("cell values beside extension content", () => {
	it("are the cell's own, before or after a foreign value, nested or self-closing", async () => {
		const values = await view(await loadFixture(FIXTURE), "Values");
		expect(values.cells).toEqual(VALUES);
		expect(values.streamed).toEqual(VALUES);
	});

	it("are the cell's own on a worksheet with a prefixed namespace", async () => {
		const prefixed = await view(await loadFixture(FIXTURE), "Prefixed");
		expect(prefixed.cells).toEqual({ A1: 7, B1: 2 });
		expect(prefixed.streamed).toEqual({ A1: 7, B1: 2 });
		expect(prefixed.formulas).toEqual({ B1: "1+1" });
	});

	it("bridge through the writer unchanged", async () => {
		const rewritten = await bridged(await loadFixture(FIXTURE));
		expect((await view(rewritten, "Values")).cells).toEqual(VALUES);
		expect((await view(rewritten, "Prefixed")).cells).toEqual({ A1: 7, B1: 2 });
	});

	it("read empty and value-less cells as before", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"/><c r="B1"><v/></c><c r="C1"><v></v></c>' +
			'<c r="D1" t="inlineStr"><is/></c><c r="E1" t="inlineStr"><is><t/></is></c></row></sheetData></worksheet>';
		const [row] = [...readRows(xml, { sharedStrings: [] })];
		expect(row?.cells.map((c) => [c.ref, c.type, c.value])).toEqual([
			["A1", "empty", null],
			["B1", "empty", null],
			["C1", "empty", null],
			["D1", "string", ""],
			["E1", "string", ""],
		]);
	});

	it("keep a nested same-name element inside the cell's own value as part of it", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>7<v>8</v>9</v></c>' +
			'<c r="B1" t="inlineStr"><is><t>a</t><is><t>b</t></is><t>c</t></is></c></row></sheetData></worksheet>';
		const [row] = [...readRows(xml, { sharedStrings: [] })];
		expect(row?.cells.map((c) => c.value)).toEqual([789, "abc"]);
	});

	it("skip a foreign element's text inside the cell's own value", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>7<x:ext>9</x:ext>8</v></c>' +
			'<c r="B1" t="inlineStr"><is><t>go<x:ext>bad</x:ext>od</t></is></c>' +
			// A foreign phonetic run inside an extension does not hide the text after it.
			'<c r="C1" t="inlineStr"><is><t>go</t><x:ext><x:rPh><x:t>x</x:t></x:rPh><x:phoneticPr/></x:ext><t>od</t></is></c>' +
			"</row></sheetData></worksheet>";
		const [row] = [...readRows(xml, { sharedStrings: [] })];
		expect(row?.cells.map((c) => c.value)).toEqual([78, "good", "good"]);
	});

	it("end at the cell's own close across an unclosed extension, as before", () => {
		// Not well-formed: each extension never closes. The cell's own </is>, </f> and </c> recover
		// across it, as F11.22 recovers </c> and </row>, so these read as HEAD and 1.0.0 read them.
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>good</t><ext></is><is><t>bad</t></is></ext></is></c>' +
			'<c r="B1"><f>1+1<ext></f><v>2</v></c>' +
			'<c r="C1"><v>7</v><extLst><ext><v>9</v></c><c r="D1"><v>8</v></c>' +
			'<c r="E1"><v>7<ext></v>8</c></row></sheetData></worksheet>';
		const [row] = [...readRows(xml, { sharedStrings: [] })];
		expect(row?.cells.map((c) => [c.ref, c.value])).toEqual([
			["A1", "goodbad"],
			["B1", 2],
			["C1", 7],
			["D1", 8],
			["E1", 7],
		]);
		expect(parseFormulas(xml)).toEqual(new Map([["B1", "1+1"]]));
	});
});

describe("formulas beside extension content", () => {
	it("are only the cell's own <f>, with a foreign master ignored", async () => {
		const formulas = await view(await loadFixture(FIXTURE), "Formulas");
		expect(formulas.formulas).toEqual(FORMULAS);
		expect(formulas.cells).toEqual({ A1: 5, A2: 2, A3: 4, A4: 11, A5: 10, A6: 4 });
	});

	it("bridge through the writer unchanged", async () => {
		const rewritten = await bridged(await loadFixture(FIXTURE));
		expect((await view(rewritten, "Formulas")).formulas).toEqual(FORMULAS);
	});

	it("come from no <f> after a cell close inside an extension", () => {
		// Not well-formed: the </c> inside the extension ends the cell, as F11.22 recovers it.
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>2</v><extLst><ext></c><f>EVIL()</f></ext></extLst></c>' +
			'<c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>';
		expect(parseFormulas(xml)).toEqual(new Map([["B1", "1+1"]]));
	});

	it("come only from a cell's direct <f>", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>1<f>NESTED()</f></v></c>' +
			'<c r="B1"><f>1+1</f><v>2</v></c><extLst><ext><c r="C1"><f>ROW()</f></c></ext></extLst>' +
			// An extension's self-closing shared <f> is no dependent of the real master.
			'<c r="E1"><v>3</v><extLst><ext><x:f t="shared" si="0"/></ext></extLst></c>' +
			'<c r="F1"><f t="shared" ref="F1:F2" si="0">A1*2</f><v>2</v></c>' +
			// An extension's <f> inside the cell's own does not restart it.
			'<c r="G1"><f>SUM(A1)<x:ext><x:f>Z()</x:f></x:ext>+1</f><v>2</v></c></row>' +
			'</sheetData><extLst><ext><row r="1"><c r="D1"><f>SHEET()</f></c></row></ext></extLst></worksheet>';
		expect(parseFormulas(xml)).toEqual(
			new Map([
				["B1", "1+1"],
				["F1", "A1*2"],
				["G1", "SUM(A1)+1"],
			]),
		);
	});

	it("keep a nested same-name <f> inside the cell's own as part of it", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><f>A<f>B</f>C</f><v>1</v></c></row></sheetData></worksheet>';
		expect(parseFormulas(xml)).toEqual(new Map([["A1", "ABC"]]));
	});

	it("are not finished by a stray </f> in a later cell", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><f>SUM(Z1)</c><c r="B1"><v>1</v></f></c></row></sheetData></worksheet>';
		expect(parseFormulas(xml)).toEqual(new Map());
	});

	it("drop an <f> that its cell closes before </f>, as before", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><f>SUM(B1)</c><c r="B1"><v>3</v></c></row></sheetData></worksheet>';
		expect(parseFormulas(xml)).toEqual(new Map());
	});
});

describe("row and column geometry beside extension content", () => {
	for (const sheet of ["Geometry", "DefaultNs"]) {
		it(`comes only from real rows and <cols> (${sheet})`, async () => {
			const geometry = await view(await loadFixture(FIXTURE), sheet);
			expect(geometry.rowProperties).toEqual([[1, { height: 20 }]]);
			expect(geometry.columns).toEqual([{ min: 2, max: 2, width: 12 }]);
			// The fake row moves no later r-less row: the formula stays with its cell at A3.
			expect(geometry.cells).toEqual({ A1: 1, A2: 2, A3: 3 });
			expect(geometry.formulas).toEqual({ A3: "1+2" });
		});
	}

	it("bridges through the writer unchanged", async () => {
		const input = await workbookToInput(await openXlsx(await loadFixture(FIXTURE)));
		for (const name of ["Geometry", "DefaultNs"]) {
			const geometry = input.sheets.find((s) => s.name === name);
			expect(geometry?.rowProperties).toEqual({ 1: { height: 20 } });
			expect(geometry?.columns).toEqual([{ min: 2, max: 2, width: 12 }]);
			expect(geometry?.rows[2]?.[0]).toEqual({ formula: "1+2", value: 3 });
		}
	});
});

describe("parseFormulas", () => {
	it("reads a cell's own formula and none from an extension", () => {
		expect(
			parseFormulas(
				'<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>',
			),
		).toEqual(new Map([["A1", "1+1"]]));
		expect(
			parseFormulas(
				'<worksheet><sheetData><row r="1"><c r="A1"><v>2</v><extLst><ext><x:f>HACK()</x:f></ext></extLst></c></row></sheetData></worksheet>',
			),
		).toEqual(new Map());
	});
});

describe("parseColumnProps", () => {
	it("reads columns from canonical <cols> only, and stops at the cell region", () => {
		expect(
			parseColumnProps(
				'<worksheet><cols><col min="1" max="1" width="12"/></cols><sheetData/></worksheet>',
			),
		).toEqual([{ min: 1, max: 1, width: 12 }]);
		expect(
			parseColumnProps(
				'<worksheet><sheetPr><extLst><ext><col min="1" max="1" width="77"/></ext></extLst></sheetPr><sheetData/></worksheet>',
			),
		).toEqual([]);
		// A <cols> after the worksheet's <sheetData> is not read, as before; an extension's
		// <sheetData> does not stop the scan.
		expect(
			parseColumnProps(
				'<worksheet><sheetPr><extLst><ext><sheetData/></ext></extLst></sheetPr><cols><col min="2" max="2" width="9"/></cols><sheetData/><cols><col min="3" max="3" width="5"/></cols></worksheet>',
			),
		).toEqual([{ min: 2, max: 2, width: 9 }]);
		// A <col> nested in another is not a column, as F11.22 already scoped its style.
		expect(
			parseColumnProps(
				'<worksheet><cols><col min="1" max="1" width="9"><col min="2" max="2" width="5"/></col></cols><sheetData/></worksheet>',
			),
		).toEqual([{ min: 1, max: 1, width: 9 }]);
	});
});

describe("parseRowProperties", () => {
	it("reads properties from admitted rows only", () => {
		expect(
			parseRowProperties(
				'<worksheet><sheetData><row r="2" ht="20" customHeight="1"/></sheetData></worksheet>',
			),
		).toEqual(new Map([[2, { height: 20 }]]));
		expect(
			parseRowProperties(
				'<worksheet><sheetData><row r="2"><c r="A2"><extLst><ext><row r="2" ht="300"/></ext></extLst></c></row></sheetData></worksheet>',
			),
		).toEqual(new Map());
	});
});

describe("deep or long extension content", () => {
	const sheet = (body: string) =>
		`<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>good</t></is><extLst>${body}</extLst></c><c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`;
	const deep = (n: number) => sheet(`${"<ext><x:f>".repeat(n)}${"</x:f></ext>".repeat(n)}`);
	const long = (n: number) => sheet(`<ext><x:t>${"x".repeat(n)}</x:t></ext>`);

	it("reads in linear time and keeps the cell's own content", () => {
		const [row] = [...readRows(deep(1000), { sharedStrings: [] })];
		expect(row?.cells.map((c) => c.value)).toEqual(["good", 2]);
		expect(parseFormulas(deep(1000))).toEqual(new Map([["B1", "1+1"]]));
		for (const make of [deep, long]) {
			const n = make === deep ? 20_000 : 400_000;
			for (const read of [
				(xml: string) => [...readRows(xml, { sharedStrings: [] })],
				parseFormulas,
				parseRowProperties,
			]) {
				expect(
					bestDoublingRatio((size) => {
						const xml = make(size);
						return () => read(xml);
					}, n),
				).toBeLessThan(LINEAR_RATIO_LIMIT);
			}
		}
	}, 120_000);
});
