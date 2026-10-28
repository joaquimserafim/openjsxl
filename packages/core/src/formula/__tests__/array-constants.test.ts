import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateWorkbook } from "../eval";
import { errorValue, isRangeView } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.64: an array constant passed to a function is an array of its elements. Data!A1:C2 is a small
// table for the lookups: 1, "x", "y" over 2, "p", "q".

const data: readonly (readonly CellInput[])[] = [
	[1, "x", "y"],
	[2, "p", "q"],
];

describe("an array constant in a function argument", () => {
	it("gives array-consuming built-ins every element", async () => {
		expect(
			await evaluateFormulas(
				[
					"SUM({1,2,3})",
					"MAX({1,5,3})",
					"AVERAGE({1,2,3,4})",
					"LARGE({1,5,3},2)",
					'COUNTA({1,"a",TRUE})',
					"SUMPRODUCT({1,2},{3,4})",
					'INDEX({"Jan","Feb","Mar"},2)',
					'MATCH("b",{"a","b","c"},0)',
					'VLOOKUP(2,{1,"a";2,"b"},2,FALSE)',
				],
				data,
			),
		).toEqual([6, 5, 2.5, 3, 3, 11, "Feb", 2, "b"]);
	});

	it("keeps its rows and columns", async () => {
		expect(
			await evaluateFormulas(
				["SUM({1,2;3,4})", "INDEX({1,2;3,4},2,1)", "ROWS({1,2;3,4})", "COLUMNS({1,2;3,4})"],
				data,
			),
		).toEqual([10, 3, 2, 2]);
	});

	it("follows Excel's rules for text and booleans in an array", async () => {
		expect(
			await evaluateFormulas(
				[
					'SUM({1,"2",TRUE})',
					'COUNT({1,"a",TRUE})',
					'MIN({2,"x",1})',
					"AND({TRUE,FALSE})",
					"OR({FALSE,TRUE})",
				],
				data,
			),
		).toEqual([1, 1, 1, false, true]);
	});

	it("reaches the function through a lazy argument or parentheses", async () => {
		expect(
			await evaluateFormulas(
				[
					"SUM(IF(TRUE,{1,2,3}))",
					"SUM(CHOOSE(2,{1,2},{3,4}))",
					"SUM(({1,2,3}))",
					"SUM({-1,+2})",
				],
				data,
			),
		).toEqual([6, 7, 6, 1]);
	});

	it("is read as its first element by a parameter that takes one value", async () => {
		// Excel shows the first element's result for an array in a single-value parameter.
		expect(
			await evaluateFormulas(
				["ABS({-1,2})", 'LEN({"ab","c"})', "VLOOKUP(2,Data!A1:C2,{2,3},FALSE)", "N({5,6})"],
				data,
			),
		).toEqual([1, 2, "p", 5]);
	});

	it("propagates an error element through an aggregate", async () => {
		expect(await evaluateFormulas(["SUM({#N/A,1})", "IFERROR(SUM({#N/A,1}),9)"], data)).toEqual(
			[errorValue("#N/A"), 9],
		);
	});

	it("reaches a caller function as a range of its elements", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[{ formula: "SIZE({1,2;3,4})", value: 0 }]] }],
			}),
		);
		const result = await evaluateWorkbook(book, {
			functions: {
				SIZE: {
					minArgs: 1,
					maxArgs: 1,
					evaluate: ([v]: readonly unknown[]) => (isRangeView(v) ? v.cellCount : -1),
				},
			},
		});
		expect(result.get("S", "A1")).toBe(4);
	});
});

describe("a formula whose value is an array constant", () => {
	it("shows its first element, at the top level or through a function (decision 5)", async () => {
		expect(
			await evaluateFormulas(
				["{1,2,3}", "{4;5}", "IF(TRUE,{1,2,3})", "IFERROR({1,2},0)"],
				data,
			),
		).toEqual([1, 4, 1, 1]);
	});

	it("holds its first element for the formulas that read it", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							[
								{ formula: "IF(TRUE,{1,2,3})", value: 0 },
								{ formula: "SUM(A1)", value: 0 },
								{ formula: "A1+1", value: 0 },
								{ formula: "COUNTA(A1)", value: 0 },
							],
						],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect(["A1", "B1", "C1", "D1"].map((ref) => result.get("S", ref))).toEqual([1, 1, 2, 1]);
	});
});

describe("array constants F11.64 leaves as they were", () => {
	it("an operator on an array constant still uses its first element", async () => {
		// Element-wise array operators are out of scope: Excel gives {2,4} and SUM 6 here.
		expect(await evaluateFormulas(["{1,2}*2", "SUM({1,2}*2)", "SUM(-{1,2})"], data)).toEqual([
			2, 2, -1,
		]);
	});

	it("a formula with rows of different lengths does not parse, so it is #VALUE!", async () => {
		expect(await evaluateFormulas(["SUM({1,2;3})", "MATCH(3,{1;2,3},0)"], data)).toEqual([
			errorValue("#VALUE!"),
			errorValue("#VALUE!"),
		]);
	});
});
