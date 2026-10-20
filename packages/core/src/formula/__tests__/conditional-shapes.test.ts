import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateWorkbook } from "../eval";
import { errorValue } from "../value";

// F11.11: strict *IFS rectangles and complete criterion pairs. Formula cells live on a separate
// sheet so whole-column inputs contain only the supplied data. F11.12's all-blank multi-range
// iteration is deliberately not pinned here; valid multi-range controls exclude blanks.
async function evaluate(formulas: readonly string[], rows: readonly (readonly CellInput[])[]) {
	const book = await openXlsx(
		await writeXlsx({
			sheets: [
				{ name: "Data", rows },
				{
					name: "Results",
					rows: formulas.map((formula) => [{ formula, value: 0 }]),
				},
			],
		}),
	);
	const result = await evaluateWorkbook(book);
	return formulas.map((_, index) => result.get("Results", `A${index + 1}`));
}

async function expectShapeErrors(formulas: readonly string[]) {
	const values = await evaluate(formulas, [
		[1, 1, 3, 1],
		[1, 1, 4, 1],
		[1, 1, 5, 1],
		[1, 1, 6, 1],
		[1, 1, 7, 1],
	]);
	for (let index = 0; index < formulas.length; index++) {
		expect(values[index], formulas[index]).toEqual(errorValue("#VALUE!"));
	}
}

describe("conditional aggregates require aligned *IFS shapes (F11.11)", () => {
	it("rejects too few arguments and incomplete criterion pairs, including trailing commas", async () => {
		await expectShapeErrors([
			"COUNTIFS()",
			"COUNTIFS(Data!A1:A2)",
			'COUNTIFS(Data!A1:A2,">0",Data!B1:B2)',
			'COUNTIFS(Data!A1:A2,">0",)',
			'COUNTIFS(Data!A1:A2,">0",Data!B1:B2,">0",Data!C1:C2)',
			"SUMIFS()",
			"SUMIFS(Data!C1:C2)",
			"SUMIFS(Data!C1:C2,Data!A1:A2)",
			'SUMIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2)',
			'SUMIFS(Data!C1:C2,Data!A1:A2,">0",)',
			'SUMIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2,">0",Data!D1:D2)',
			"AVERAGEIFS()",
			"AVERAGEIFS(Data!C1:C2)",
			"AVERAGEIFS(Data!C1:C2,Data!A1:A2)",
			'AVERAGEIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2)',
			'AVERAGEIFS(Data!C1:C2,Data!A1:A2,">0",)',
			'AVERAGEIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2,">0",Data!D1:D2)',
		]);
	});

	it("rejects both audited counterexamples through the public workbook evaluator", async () => {
		await expectShapeErrors([
			'COUNTIFS(Data!A1:A2,">0",Data!B1:B3,">0")',
			'SUMIFS(Data!C1:C3,Data!A1:A2,">0",Data!B1:B3,">0")',
		]);
	});

	it("checks every criteria range against the required height and width", async () => {
		await expectShapeErrors([
			'COUNTIFS(Data!A1:A3,">0",Data!B1:B2,">0")',
			'COUNTIFS(Data!A1:A2,">0",Data!B1:B2,">0",Data!C1:C3,">0")',
			'COUNTIFS(Data!A1:B1,">0",Data!A2:C2,">0")',
			'SUMIFS(Data!C1:C3,Data!A1:A3,">0",Data!B1:B2,">0")',
			'SUMIFS(Data!A1:C1,Data!A2:C2,">0",Data!A3:B3,">0")',
			'AVERAGEIFS(Data!C1:C3,Data!A1:A2,">0",Data!B1:B3,">0")',
			'AVERAGEIFS(Data!A1:C1,Data!A2:B2,">0")',
		]);
	});

	it("compares both dimensions, not only the number of positions", async () => {
		await expectShapeErrors([
			'COUNTIFS(Data!A1:A2,">0",Data!B1:C1,">0")',
			'COUNTIFS(Data!A1:B2,">0",Data!C1:C4,">0")',
			'SUMIFS(Data!A1:B2,Data!C1:C4,">0")',
			'SUMIFS(Data!A1:B2,Data!A3:B4,">0",Data!C1:C4,">0")',
			'AVERAGEIFS(Data!C1:C4,Data!A1:B2,">0")',
		]);
	});

	it("treats scalars and single-cell references as 1x1 when rejecting larger ranges", async () => {
		await expectShapeErrors([
			'COUNTIFS(Data!A1,">0",Data!B1:B2,">0")',
			'COUNTIFS(Data!A1:A2,">0",1,">0")',
			'SUMIFS(7,Data!A1:A2,">0")',
			'SUMIFS(Data!B1:B2,Data!A1,">0")',
			'AVERAGEIFS(Data!B1,Data!A1:A2,">0")',
			'AVERAGEIFS(Data!B1:B2,1,">0")',
		]);
	});

	it("requires equal shapes for single-pair SUMIFS and AVERAGEIFS", async () => {
		await expectShapeErrors([
			'SUMIFS(Data!B1:B5,Data!A1:A3,">0")',
			'AVERAGEIFS(Data!B1:B5,Data!A1:A3,">0")',
			'SUMIFS(Data!B1:B2,Data!A1:A3,">0")',
			'AVERAGEIFS(Data!B1:B2,Data!A1:A3,">0")',
		]);
	});

	it("rejects shape mismatches even when every range is empty or spans a whole column", async () => {
		await expectShapeErrors([
			'COUNTIFS(Data!X1:X2,">0",Data!Y1:Y3,">0")',
			'SUMIFS(Data!Z1:Z3,Data!X1:X2,">0")',
			'AVERAGEIFS(Data!Z1:Z3,Data!X1:X2,">0")',
			'COUNTIFS(Data!A:A,">0",Data!B1:B3,">0")',
			'SUMIFS(Data!B:B,Data!A1:A3,">0")',
			'AVERAGEIFS(Data!B1:B3,Data!A:A,">0")',
		]);
	});

	it("aligns equal 2-D shapes by offset even when their origins differ", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:B2,">0",Data!D3:E4,">0")',
				'SUMIFS(Data!G5:H6,Data!A1:B2,">0",Data!D3:E4,">0")',
				'AVERAGEIFS(Data!G5:H6,Data!A1:B2,">0",Data!D3:E4,">0")',
			],
			[
				[1, 0],
				[1, 1],
				[null, null, null, 1, 1],
				[null, null, null, 0, 1],
				[null, null, null, null, null, null, 10, 20],
				[null, null, null, null, null, null, 30, 40],
			],
		);
		expect(values).toEqual([2, 50, 25]);
	});

	it("accepts equivalent scalar, cell and one-cell-range shapes", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1,">0",Data!B1:B1,">0")',
				'COUNTIFS(1,">0",Data!B1,">0")',
				'SUMIFS(Data!C1,Data!A1:A1,">0",Data!B1,">0")',
				'SUMIFS(7,1,">0")',
				'AVERAGEIFS(Data!C1:C1,Data!A1,">0",Data!B1:B1,">0")',
				'AVERAGEIFS(7,1,">0")',
			],
			[[1, 1, 3]],
		);
		expect(values).toEqual([1, 1, 3, 7, 3, 7]);
	});

	it("preserves single-pair results and SUMIF/AVERAGEIF top-left reshaping", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A3,">0")',
				'SUMIFS(Data!B1:B3,Data!A1:A3,">0")',
				'AVERAGEIFS(Data!B1:B3,Data!A1:A3,">0")',
				'SUMIF(Data!A1:A3,">0",Data!B1:B5)',
				'AVERAGEIF(Data!A1:A3,">0",Data!B1:B5)',
			],
			[
				[1, 10],
				[0, 20],
				[1, 30],
				[null, 40],
				[null, 50],
			],
		);
		expect(values).toEqual([2, 40, 20, 40, 20]);
	});

	it("keeps explicit empty criterion slots when the complete pair count is valid", async () => {
		// Empty AST arguments already evaluate to null (a blank criterion). This task rejects
		// incomplete pairs, not valid pairs whose criterion was explicitly left empty.
		const values = await evaluate(
			[
				"COUNTIFS(Data!A1:A3,)",
				"SUMIFS(Data!B1:B3,Data!A1:A3,)",
				"AVERAGEIFS(Data!B1:B3,Data!A1:A3,)",
				'COUNTIFS(Data!A1:A3,,Data!B1:B3,">0")',
				'SUMIFS(Data!B1:B3,Data!A1:A3,,Data!B1:B3,">0")',
			],
			[
				[0, 10],
				[1, 20],
				[null, 30],
			],
		);
		expect(values).toEqual([1, 30, 30, 1, 30]);
	});

	it("preserves aligned error propagation and no-match results", async () => {
		const values = await evaluate(
			[
				"COUNTIFS(Data!A1:A2,1/0)",
				"SUMIFS(Data!B1:B2,Data!A1:A2,1/0)",
				"AVERAGEIFS(Data!B1:B2,Data!A1:A2,1/0)",
				'SUMIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2,">0")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,">0",Data!B1:B2,">0")',
				'COUNTIFS(Data!A1:A2,">99",Data!B1:B2,">0")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,">99",Data!B1:B2,">0")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,">99",Data!B1:B2,">0")',
			],
			[
				[1, 1, { formula: "1/0", value: 0 }],
				[0, 1, 4],
			],
		);
		const div0 = errorValue("#DIV/0!");
		expect(values).toEqual([div0, div0, div0, div0, div0, 0, 0, div0]);
	});
});
