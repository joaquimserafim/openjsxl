import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx, streamSheetRows } from "../workbook";
import { rewriteParts } from "./shared-part-fixtures";

// F11.38 no-change cases. Excel 2010 (x14) worksheet extensions after </sheetData> carry their own
// <xm:f> and <xm:sqref>; they supply no formula, row or column, and every accessor reads as before,
// with a self-closing <sheetData/> too. The F11.32 macro sheets keep 1.0.0's formulas and geometry.

async function view(bytes: Uint8Array, sheet: string) {
	const ws = (await openXlsx(bytes)).sheet(sheet);
	const cells: Record<string, unknown> = {};
	const formulas: Record<string, string> = {};
	for await (const row of ws.rows()) {
		for (const cell of row.cells) {
			cells[cell.ref] = cell.value;
			const formula = ws.formula(cell.ref);
			if (formula !== undefined) formulas[cell.ref] = formula;
		}
	}
	const streamed: Record<string, unknown> = {};
	for await (const row of streamSheetRows(bytes, sheet)) {
		for (const cell of row.cells) streamed[cell.ref] = cell.value;
	}
	return {
		cells,
		streamed,
		formulas,
		rowProperties: [...ws.rowProperties],
		columns: ws.columns,
		dataValidations: ws.dataValidations,
		conditionalFormatting: ws.conditionalFormatting,
	};
}

describe("x14 worksheet extensions after </sheetData>", () => {
	it("leave the cells, formulas and geometry as they are", async () => {
		const data = await view(await loadFixture("x14-worksheet-extlst.xlsx"), "Data");
		const cells = { A1: 10, B1: 20, C1: 30, D1: 60, A2: 5, B2: "Yes", A3: 15 };
		expect(data.cells).toEqual(cells);
		expect(data.streamed).toEqual(cells);
		expect(data.formulas).toEqual({ D1: "SUM(A1:C1)" });
		expect(data.rowProperties).toEqual([[1, { height: 18 }]]);
		expect(data.columns).toEqual([{ min: 1, max: 1, width: 14 }]);
		expect(data.dataValidations).toEqual([]);
		expect(data.conditionalFormatting).toEqual([
			{
				sqref: ["A1:A3"],
				rules: [
					{
						type: "dataBar",
						priority: 1,
						cfvo: [{ type: "min" }, { type: "max" }],
						color: { rgb: "FF638EC6" },
					},
				],
			},
		]);
	});

	it("leave a sheet with a self-closing <sheetData/> empty", async () => {
		const empty = await view(await loadFixture("x14-worksheet-extlst.xlsx"), "Empty");
		expect(empty).toEqual({
			cells: {},
			streamed: {},
			formulas: {},
			rowProperties: [],
			columns: [],
			dataValidations: [],
			conditionalFormatting: [],
		});
	});
});

describe("Excel 4.0 macro sheets", () => {
	it("keep 1.0.0's formulas, row properties and columns", async () => {
		const bytes = await loadFixture("macrosheet-xlm.xlsm");
		const macro1 = await view(bytes, "Macro1");
		const macro2 = await view(bytes, "Macro2");
		expect([macro1.formulas, macro1.rowProperties, macro1.columns]).toEqual([
			{ A5: "RETURN(A1)" },
			[],
			[],
		]);
		expect([macro2.formulas, macro2.rowProperties, macro2.columns]).toEqual([
			{ A2: "RETURN(A1*2)" },
			[],
			[],
		]);
	});

	it("keep 1.0.0's row properties and columns when they declare them", async () => {
		const bytes = await rewriteParts(await loadFixture("macrosheet-xlm.xlsm"), {
			"xl/macrosheets/sheet1.xml": (xml) =>
				xml
					.replace(
						"<sheetData>",
						'<cols><col min="1" max="1" width="20" customWidth="1"/></cols><sheetData>',
					)
					.replace('<row r="1">', '<row r="1" ht="25" customHeight="1">'),
		});
		const macro1 = await view(bytes, "Macro1");
		expect([macro1.formulas, macro1.rowProperties, macro1.columns]).toEqual([
			{ A5: "RETURN(A1)" },
			[[1, { height: 25 }]],
			[{ min: 1, max: 1, width: 20 }],
		]);
	});
});
