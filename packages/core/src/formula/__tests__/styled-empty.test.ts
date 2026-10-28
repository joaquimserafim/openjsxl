import { describe, expect, it } from "vitest";
import { rewriteParts } from "../../reader/__tests__/shared-part-fixtures";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.63: a cell that carries only a style is empty to every formula, as in Excel. Data!A1:A5 holds
// 1, 3, a styled empty cell, 7 and 9, with labels in B; Data!E1:I2 holds the same laid out across.

const bold: CellInput = { value: null, style: { font: { bold: true } } };

function rows(blank: CellInput): readonly (readonly CellInput[])[] {
	return [
		[1, "one", null, null, 1, 3, blank, 7, 9],
		[3, "three", null, null, "one", "three", "blank-col", "seven", "nine"],
		[blank, "blank-row"],
		[7, "seven"],
		[9, "nine"],
	];
}

describe("a styled empty cell", () => {
	it("is not counted by COUNTA", async () => {
		expect(await evaluateFormulas(["COUNTA(Data!A1:A5)"], rows(bold))).toEqual([4]);
	});

	it("is skipped by an approximate VLOOKUP, HLOOKUP and MATCH", async () => {
		expect(
			await evaluateFormulas(
				["VLOOKUP(6,Data!A1:B5,2)", "HLOOKUP(6,Data!E1:I2,2)", "MATCH(6,Data!A1:A5,1)"],
				rows(bold),
			),
		).toEqual(["three", "three", 2]);
	});

	it("is not found by MATCH of 0 or an empty string", async () => {
		expect(
			await evaluateFormulas(
				["MATCH(0,Data!A1:A5,0)", 'MATCH("",Data!A1:A5,0)', "MATCH(0,Data!E1:I1,0)"],
				rows(bold),
			),
		).toEqual([errorValue("#N/A"), errorValue("#N/A"), errorValue("#N/A")]);
	});

	it("is still blank to ISBLANK, COUNTBLANK and a blank criterion", async () => {
		expect(
			await evaluateFormulas(
				[
					"ISBLANK(Data!A3)",
					"COUNTBLANK(Data!A1:A5)",
					'COUNTIF(Data!A1:A5,"")',
					'COUNTIF(Data!A1:A5,"<>")',
					"COUNT(Data!A1:A5)",
					'Data!A3&"x"',
				],
				rows(bold),
			),
		).toEqual([true, 1, 1, 4, 4, "x"]);
	});

	it("gives every formula the result an unformatted empty cell gives", async () => {
		const formulas = [
			"COUNTA(Data!A1:A5)",
			"COUNTA(Data!E1:I1)",
			"VLOOKUP(6,Data!A1:B5,2)",
			"VLOOKUP(6,Data!A1:B5,2,TRUE)",
			"HLOOKUP(6,Data!E1:I2,2)",
			"MATCH(6,Data!A1:A5,1)",
			"MATCH(0,Data!A1:A5,0)",
			'MATCH("",Data!A1:A5,0)',
			"MATCH(0,Data!E1:I1,0)",
			"COUNT(Data!A1:A5)",
			"COUNTBLANK(Data!A1:A5)",
			'COUNTIF(Data!A1:A5,"")',
			'COUNTIF(Data!A1:A5,"<>")',
			"SUM(Data!A1:A5)",
			"AVERAGE(Data!A1:A5)",
			"MIN(Data!A1:A5)",
			"ROWS(Data!A1:A5)",
			"ISBLANK(Data!A3)",
			'INDEX(Data!A1:A5,3)&"x"',
			"Data!A3",
		];
		const plain = await evaluateFormulas(formulas, rows(null));
		expect(await evaluateFormulas(formulas, rows(bold))).toEqual(plain);
		expect(plain.slice(0, 3)).toEqual([4, 4, "three"]);
	});

	it("gives TEXTJOIN with ignore_empty FALSE the result an absent cell gives", async () => {
		// Both get a slot, as Excel keeps one for every empty cell (F11.97).
		const formula = ['TEXTJOIN(",",FALSE,Data!A1:A3)'];
		const absent = await evaluateFormulas(formula, [["a"], [], ["c"]]);
		expect(await evaluateFormulas(formula, [["a"], [bold], ["c"]])).toEqual(absent);
	});

	it("is left out, while a styled cell with a value or a formula still counts", async () => {
		const styled: readonly (readonly CellInput[])[] = [
			[{ value: 5, style: { font: { bold: true } } }],
			[{ formula: '""', style: { font: { bold: true } } }],
			[bold],
		];
		expect(await evaluateFormulas(["COUNTA(Data!A1:A3)", "COUNT(Data!A1:A3)"], styled)).toEqual(
			[2, 1],
		);
	});
});

describe("what F11.63 leaves as it was", () => {
	it("the styled empty cell itself evaluates as empty, and a formula over it as 0", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "Data", rows: [[bold, { formula: "A1", value: 0 }]] }],
			}),
		);
		expect(await evaluateCell(book, "Data", "A1")).toBeNull();
		expect(await evaluateCell(book, "Data", "B1")).toBe(0);
	});
});

describe("a cell the reader reads as empty", () => {
	// Data!A1 is 5 and A2 a styled empty cell (style 1); each case rewrites Data's sheetData.
	async function evaluated(sheetData: string) {
		const base = await writeXlsx({
			sheets: [
				{ name: "Data", rows: [[5], [bold]] },
				{
					name: "Results",
					rows: [
						[{ formula: "COUNTA(Data!A1:A2)", value: 0 }],
						[{ formula: 'MATCH("",Data!A1:A2,0)', value: 0 }],
						[{ formula: 'Data!A1&"x"', value: 0 }],
					],
				},
			],
		});
		const book = await openXlsx(
			await rewriteParts(base, {
				"xl/worksheets/sheet1.xml": (xml) =>
					xml.replace(
						/<sheetData>[\s\S]*<\/sheetData>/,
						`<sheetData>${sheetData}</sheetData>`,
					),
			}),
		);
		const result = await evaluateWorkbook(book);
		return {
			cell: book.sheet("Data").cell("A1"),
			results: ["A1", "A2", "A3"].map((ref) => result.get("Results", ref)),
		};
	}

	it("lets a later empty copy of a ref win, as Worksheet.cell() does", async () => {
		const later = await evaluated(
			'<row r="1"><c r="A1"><v>5</v></c><c r="A1" s="1"/></row><row r="2"><c r="A2"><v>2</v></c></row>',
		);
		expect(later.cell.type).toBe("empty");
		expect(later.results).toEqual([1, errorValue("#N/A"), "x"]);
		const split = await evaluated(
			'<row r="1"><c r="A1"><v>5</v></c></row><row r="1"><c r="A1" s="1"/></row><row r="2"><c r="A2"><v>2</v></c></row>',
		);
		expect(split.cell.type).toBe("empty");
		expect(split.results).toEqual([1, errorValue("#N/A"), "x"]);
		const earlier = await evaluated(
			'<row r="1"><c r="A1" s="1"/><c r="A1"><v>5</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>',
		);
		expect(earlier.cell).toMatchObject({ type: "number", value: 5 });
		expect(earlier.results).toEqual([2, errorValue("#N/A"), "5x"]);
	});

	it("leaves out a value the reader degrades to empty", async () => {
		// Text in a t="d" cell that is not a date reads as empty (F11.60).
		const degraded = await evaluated(
			'<row r="1"><c r="A1" t="d"><v>not a date</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>',
		);
		expect(degraded.cell.type).toBe("empty");
		expect(degraded.results).toEqual([1, errorValue("#N/A"), "x"]);
	});
});
