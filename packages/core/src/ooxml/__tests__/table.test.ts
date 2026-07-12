import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { MAX_TABLE_NAME_LEN, parseTable } from "../table";

// F9.1 — reading table parts (xl/tables/tableN.xml). Unit-pins the parser on small XML literals plus
// a verbatim read of the committed Excel-authored `inventory-table.xlsx` fixture.

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

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
