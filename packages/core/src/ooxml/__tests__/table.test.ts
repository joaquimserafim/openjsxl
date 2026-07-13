import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { MAX_TABLE_NAME_LEN, normalizeTableName, parseTable, tableNameProblem } from "../table";

// F9.1 — reading table parts (xl/tables/tableN.xml). Unit-pins the parser on small XML literals plus
// a verbatim read of the committed Excel-authored `inventory-table.xlsx` fixture.

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

// F9.5 — the single-sourced table-name rule that BOTH the tolerant reader (normalizes) and the strict
// writer (rejects) consume. Pinning it here keeps the two sides provably in lockstep.
describe("tableNameProblem — the shared table-name legality rule", () => {
	it("accepts legal identifiers", () => {
		for (const ok of ["Table1", "Sales", "_hidden", "T", "ABCD1", "\\weird"]) {
			expect(tableNameProblem(ok)).toBeUndefined();
		}
	});

	it("classifies each illegal name by its FIRST broken rule (writer's check order)", () => {
		expect(tableNameProblem("")).toBe("empty");
		expect(tableNameProblem("a".repeat(MAX_TABLE_NAME_LEN + 1))).toBe("too-long");
		expect(tableNameProblem("a\x01b")).toBe("not-xml-safe");
		expect(tableNameProblem("My Table")).toBe("whitespace");
		expect(tableNameProblem("1abc")).toBe("bad-start");
		expect(tableNameProblem("データ")).toBe("bad-start"); // rule is ASCII-start only (pre-existing)
		expect(tableNameProblem("A1")).toBe("cell-ref");
		expect(tableNameProblem("ABC12")).toBe("cell-ref");
		expect(tableNameProblem("C")).toBe("cell-ref");
		expect(tableNameProblem("R")).toBe("cell-ref");
	});
});

describe("normalizeTableName — repair an illegal name into a legal identifier (F9.5)", () => {
	it("returns an already-legal name UNCHANGED (clean files stay byte-identical)", () => {
		for (const ok of ["Table1", "Sales", "_hidden", "\\weird", "T"]) {
			expect(normalizeTableName(ok)).toBe(ok);
		}
	});

	it("repairs each illegal category to something legal", () => {
		expect(normalizeTableName("My Table")).toBe("My_Table"); // whitespace → _
		expect(normalizeTableName("A1")).toBe("A1_"); // cell-ref broken with trailing _
		expect(normalizeTableName("1abc")).toBe("_1abc"); // bad start prefixed
		expect(normalizeTableName("データ")).toBe("_データ"); // non-ASCII start prefixed, body kept
		expect(normalizeTableName("C")).toBe("C_"); // reserved bare C
		expect(normalizeTableName("")).toBe("_"); // empty → a legal placeholder
	});

	it("ALWAYS produces a legal name (the result never has a problem)", () => {
		for (const bad of [
			"",
			" ",
			"\t\n",
			"A1",
			"C",
			"R",
			"1",
			"9x",
			"a b c",
			"=SUM",
			"\x01\x02",
			"\uD800",
		]) {
			expect(tableNameProblem(normalizeTableName(bad))).toBeUndefined();
		}
	});

	it("clamps an over-long name to the shared bound and keeps it legal", () => {
		const out = normalizeTableName("T".repeat(MAX_TABLE_NAME_LEN + 50));
		expect(out.length).toBe(MAX_TABLE_NAME_LEN);
		expect(tableNameProblem(out)).toBeUndefined();
	});
});

describe("parseTable — F9.5 shape degradation (reader returns writer-legal tables)", () => {
	it("normalizes an illegal display name (normalize-and-keep)", () => {
		const t = parseTable(
			`<table ${NS} name="My Table" displayName="My Table" ref="A1:B2"><tableColumns count="2"><tableColumn name="A"/><tableColumn name="B"/></tableColumns></table>`,
		);
		expect(t?.name).toBe("My_Table");
		expect(tableNameProblem(t?.name ?? "x")).toBeUndefined();
	});

	it("clears columns when their count doesn't match the ref width (writer then derives)", () => {
		const t = parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:C1"><tableColumns count="2"><tableColumn name="X"/><tableColumn name="Y"/></tableColumns></table>`,
		);
		expect(t?.columns).toEqual([]);
	});

	it("clears a totals row on a single-row ref", () => {
		const t = parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:B1" totalsRowCount="1"/>`,
		);
		expect(t?.totalsRow).toBe(false);
	});

	it("leaves a legal table untouched (matching count, multi-row totals)", () => {
		const t = parseTable(
			`<table ${NS} name="Good" displayName="Good" ref="A1:B3" totalsRowCount="1"><tableColumns count="2"><tableColumn name="X"/><tableColumn name="Y"/></tableColumns></table>`,
		);
		expect(t?.name).toBe("Good");
		expect(t?.columns).toEqual([{ name: "X" }, { name: "Y" }]);
		expect(t?.totalsRow).toBe(true);
	});
});

describe("parseTable — units", () => {
	it("reads name (from displayName), ref, columns, and the default header/totals flags", () => {
		const t = parseTable(
			`<table ${NS} id="1" name="Sales" displayName="SalesTable" ref="A1:B3">` +
				'<tableColumns count="2"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Total"/></tableColumns>' +
				"</table>",
		);
		expect(t).toEqual({
			name: "SalesTable", // displayName wins over name
			ref: "A1:B3",
			columns: [{ name: "Region" }, { name: "Total" }],
			headerRow: true, // headerRowCount absent → on
			totalsRow: false, // totalsRowCount absent → off
		});
	});

	it("reads the style-info banding and the totals/header counts", () => {
		const t = parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:A9" headerRowCount="0" totalsRowCount="1">` +
				'<tableColumns count="1"><tableColumn id="1" name="X"/></tableColumns>' +
				'<tableStyleInfo name="TableStyleLight1" showFirstColumn="1" showRowStripes="0"/>' +
				"</table>",
		);
		expect(t?.headerRow).toBe(false);
		expect(t?.totalsRow).toBe(true);
		expect(t?.style).toEqual({
			name: "TableStyleLight1",
			showFirstColumn: true,
			showRowStripes: false,
		});
	});

	it("carries a column's totals-row label/function and formula child elements verbatim", () => {
		const t = parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:B4" totalsRowCount="1"><tableColumns count="2">` +
				'<tableColumn id="1" name="Item" totalsRowLabel="Total"/>' +
				'<tableColumn id="2" name="Amt" totalsRowFunction="sum">' +
				"<totalsRowFormula>SUBTOTAL(109,Amt)</totalsRowFormula>" +
				'<calculatedColumnFormula>[Item]&amp;"!"</calculatedColumnFormula>' +
				"</tableColumn></tableColumns></table>",
		);
		expect(t?.columns).toEqual([
			{ name: "Item", totalsRowLabel: "Total" },
			{
				name: "Amt",
				totalsRowFunction: "sum",
				totalsRowFormula: "SUBTOTAL(109,Amt)",
				calculatedColumnFormula: '[Item]&"!"', // entities decoded
			},
		]);
	});

	it("returns undefined when name or ref is missing (the reader drops it)", () => {
		expect(parseTable(`<table ${NS} ref="A1:B2"/>`)).toBeUndefined(); // no name
		expect(parseTable(`<table ${NS} name="T" displayName="T"/>`)).toBeUndefined(); // no ref
		expect(parseTable("<not-a-table/>")).toBeUndefined();
	});

	it("clamps an over-long display name to the shared bound", () => {
		const long = "T".repeat(MAX_TABLE_NAME_LEN + 50);
		const t = parseTable(`<table ${NS} name="${long}" displayName="${long}" ref="A1:A1"/>`);
		expect(t?.name.length).toBe(MAX_TABLE_NAME_LEN);
	});

	// F9.3 retrofit: resolve a table's/column's *DxfId to an inline DxfStyle, and DEGRADE a
	// malformed/out-of-range index to no highlight (review regression — no phantom dxf[0]).
	it("resolves a valid *DxfId to an inline DxfStyle, and drops a malformed/out-of-range one", () => {
		const dxfs = [{ font: { bold: true } }, { fill: { bgColor: { rgb: "FFFFC7CE" } } }];
		const valid = parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:A2" headerRowDxfId="0"><tableColumns count="1"><tableColumn name="H" dataDxfId="1"/></tableColumns></table>`,
			dxfs,
		);
		expect(valid?.headerRowStyle).toEqual({ font: { bold: true } });
		expect(valid?.columns[0]?.dataStyle).toEqual({ fill: { bgColor: { rgb: "FFFFC7CE" } } });

		for (const bad of ["", " 0 ", "1e2", "0x1", "9"]) {
			const t = parseTable(
				`<table ${NS} name="T" displayName="T" ref="A1:A2" headerRowDxfId="${bad}"/>`,
				dxfs,
			);
			expect(t?.headerRowStyle).toBeUndefined(); // no phantom highlight
		}
		// With no dxfs table at all, a present id still resolves to nothing.
		expect(
			parseTable(`<table ${NS} name="T" displayName="T" ref="A1:A2" headerRowDxfId="0"/>`)
				?.headerRowStyle,
		).toBeUndefined();
	});
});

describe("tables — verbatim read of a real Excel fixture", () => {
	it("reads inventory-table.xlsx's table into the shared model", async () => {
		const book = await openXlsx(await loadFixture("inventory-table.xlsx"));
		const tables = book.sheet("Sheet1").tables;
		expect(tables).toEqual([
			{
				name: "Table1",
				ref: "A1:C5",
				columns: [{ name: "Item" }, { name: "Type" }, { name: "Quantity" }],
				headerRow: true,
				totalsRow: false,
				style: {
					name: "TableStyleMedium9",
					showFirstColumn: false,
					showLastColumn: false,
					showRowStripes: true,
					showColumnStripes: false,
				},
			},
		]);
	});

	it("reads openpyxl-tables.xlsx (the openpyxl-producer variant)", async () => {
		const book = await openXlsx(await loadFixture("openpyxl-tables.xlsx"));
		expect(book.sheet("People").tables).toEqual([
			{
				name: "People",
				ref: "A1:C4",
				columns: [{ name: "Name" }, { name: "Age" }, { name: "City" }],
				headerRow: true,
				totalsRow: false,
				style: {
					name: "TableStyleMedium2",
					showRowStripes: true,
					showColumnStripes: false,
				},
			},
		]);
	});
});
