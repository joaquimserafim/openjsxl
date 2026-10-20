import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";

// F11.12: implicit blanks participate in aligned multi-criterion ranges. Keep result formulas on
// another sheet so sparse/whole-column input ranges contain only the supplied data.
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

describe("conditional aggregates include sparse blank positions (F11.12)", () => {
	it("counts fully implicit blanks and aggregates their populated companion values", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A2,"",Data!B1:B2,"")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,"")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"")',
			],
			[
				[null, null, 3],
				[null, null, 4],
			],
		);
		expect(values).toEqual([2, 7, 3.5, 7, 3.5]);
	});

	it("counts disjoint and overlapping populated offsets only once", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A7,"",Data!B1:B7,"")',
				'COUNTIFS(Data!A1:A7,"",Data!B1:B7,"",Data!A1:A7,"")',
				'SUMIFS(Data!C1:C7,Data!A1:A7,"",Data!B1:B7,"")',
				'AVERAGEIFS(Data!C1:C7,Data!A1:A7,"",Data!B1:B7,"")',
				'COUNTIFS(Data!A1:A7,"<>x",Data!B1:B7,"<>x")',
				'SUMIFS(Data!C1:C7,Data!A1:A7,"<>x",Data!B1:B7,"<>x")',
				'AVERAGEIFS(Data!C1:C7,Data!A1:A7,"<>x",Data!B1:B7,"<>x")',
			],
			[
				["x", null, 10],
				[null, "x", 20],
				["x", "x", 30],
				["", null, 40],
				[null, "", 50],
				[null, null, 60],
				["y", null, 70],
			],
		);
		expect(values).toEqual([3, 3, 150, 50, 4, 220, 55]);
	});

	it("distinguishes populated formulas returning empty strings from implicit blanks without double counting", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A4,"",Data!B1:B4,"")',
				'SUMIFS(Data!C1:C4,Data!A1:A4,"",Data!B1:B4,"")',
				'AVERAGEIFS(Data!C1:C4,Data!A1:A4,"",Data!B1:B4,"")',
			],
			[
				[{ formula: '""', value: "" }, null, 3],
				[null, { formula: 'IF(TRUE,"","x")', value: "" }, 4],
				["", "", 5],
				[{ formula: 'IF(FALSE,"","x")', value: "x" }, null, 6],
			],
		);
		expect(values).toEqual([3, 12, 4]);
	});

	it("retains a blank-excluding companion driver regardless of criterion order", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A4,"",Data!B1:B4,">0")',
				'COUNTIFS(Data!B1:B4,">0",Data!A1:A4,"")',
				'COUNTIFS(Data!A1:A4,"",Data!D1:D4,"",Data!B1:B4,">0")',
				'SUMIFS(Data!C1:C4,Data!A1:A4,"",Data!B1:B4,">0")',
				'AVERAGEIFS(Data!C1:C4,Data!B1:B4,">0",Data!A1:A4,"")',
			],
			[
				[null, 2, 3],
				["", 0, 4],
				["x", 2, 5],
				[null, 3, 6],
			],
		);
		expect(values).toEqual([2, 2, 2, 9, 4.5]);
	});

	it("aggregates only numeric companion values, counting zero in the average", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A6,"",Data!B1:B6,"")',
				'SUMIFS(Data!C1:C6,Data!A1:A6,"",Data!B1:B6,"")',
				'AVERAGEIFS(Data!C1:C6,Data!A1:A6,"",Data!B1:B6,"")',
			],
			[
				[null, null, 0],
				[null, null, 2],
				[null, null, "5"],
				[null, null, true],
				[null, null, null],
				[null, null, 4],
			],
		);
		expect(values).toEqual([6, 6, 2]);
	});

	it("excludes criterion errors and ignores value errors at nonmatching positions", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A4,"",Data!B1:B4,"")',
				'COUNTIFS(Data!A1:A4,"<>x",Data!B1:B4,"<>x")',
				'SUMIFS(Data!C1:C4,Data!A1:A4,"",Data!B1:B4,"")',
				'AVERAGEIFS(Data!C1:C4,Data!A1:A4,"<>x",Data!B1:B4,"<>x")',
			],
			[
				[{ formula: "1/0", value: 0 }, null, 10],
				[null, { formula: "1/0", value: 0 }, 20],
				[null, null, 30],
				["x", null, { formula: "1/0", value: 0 }],
			],
		);
		expect(values).toEqual([1, 1, 30, 30]);
	});

	it("propagates matching value errors even when every criterion cell is implicit", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A2,"",Data!B1:B2,"")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,"")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"")',
			],
			[
				[null, null, { formula: "1/0", value: 0 }],
				[null, null, 4],
			],
		);
		const div0 = errorValue("#DIV/0!");
		expect(values).toEqual([2, div0, div0, div0, div0]);
	});

	it("aligns sparse 2-D ranges with different origins", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:B2,"",Data!D3:E4,"")',
				'SUMIFS(Data!G5:H6,Data!A1:B2,"",Data!D3:E4,"")',
				'AVERAGEIFS(Data!G5:H6,Data!A1:B2,"",Data!D3:E4,"")',
			],
			[
				[null, "x"],
				["", null],
				[null, null, null, null, null],
				[null, null, null, "x", null],
				[null, null, null, null, null, null, 1, 2],
				[null, null, null, null, null, null, 3, 4],
			],
		);
		expect(values).toEqual([2, 5, 2.5]);
	});

	it("handles one-cell ranges and scalar blank arguments as the same single offset", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1,"",Data!B1:B1,"")',
				'SUMIFS(Data!C1,Data!A1:A1,"",Data!B1,"")',
				'AVERAGEIFS(Data!C1:C1,Data!A1,"",Data!B1:B1,"")',
				'COUNTIFS("","","","")',
				'SUMIFS(5,"","","","")',
				'AVERAGEIFS(5,"","","","")',
				'SUMIFS(5,Data!A1:A1,"")',
				'AVERAGEIFS(5,Data!A1:A1,"")',
			],
			[[null, null, 3]],
		);
		expect(values).toEqual([1, 3, 3, 1, 5, 5, 5, 5]);
	});

	it("returns no-match results when populated criteria exclude every position", async () => {
		const values = await evaluate(
			[
				'COUNTIFS(Data!A1:A2,"",Data!B1:B2,"")',
				'SUMIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
				'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
			],
			[
				["x", null, { formula: "1/0", value: 0 }],
				[null, "x", { formula: "1/0", value: 0 }],
			],
		);
		expect(values).toEqual([0, 0, errorValue("#DIV/0!")]);
	});

	it("agrees with the single-pair path when the same blank criterion is repeated", async () => {
		const values = await evaluate(
			[
				'COUNTIF(Data!A1:A4,"")',
				'COUNTIFS(Data!A1:A4,"")',
				'COUNTIFS(Data!A1:A4,"",Data!A1:A4,"")',
				'SUMIF(Data!A1:A4,"",Data!B1:B4)',
				'SUMIFS(Data!B1:B4,Data!A1:A4,"")',
				'SUMIFS(Data!B1:B4,Data!A1:A4,"",Data!A1:A4,"")',
				'AVERAGEIF(Data!A1:A4,"",Data!B1:B4)',
				'AVERAGEIFS(Data!B1:B4,Data!A1:A4,"")',
				'AVERAGEIFS(Data!B1:B4,Data!A1:A4,"",Data!A1:A4,"")',
			],
			[
				[null, 3],
				["x", 4],
				["", 5],
				[null, null],
			],
		);
		expect(values).toEqual([3, 3, 3, 8, 8, 8, 4, 4, 4]);
	});

	it("counts a completely empty whole-column pair without materializing its cells", async () => {
		const values = await evaluate(['COUNTIFS(Data!A:A,"",Data!B:B,"")'], []);
		expect(values).toEqual([1_048_576]);
	});

	it("does not evaluate an unselected volatile value for single or multiple criteria", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "Data",
						rows: [
							["x", null, { formula: "NOW()", value: 0 }],
							[null, null, 4],
						],
					},
					{
						name: "Results",
						rows: [
							[
								{ formula: 'SUMIFS(Data!C1:C2,Data!A1:A2,"")', value: 0 },
								{ formula: 'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"")', value: 0 },
								{
									formula: 'SUMIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
									value: 0,
								},
								{
									formula: 'AVERAGEIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
									value: 0,
								},
							],
						],
					},
				],
			}),
		);
		// No clock is configured. Visiting the excluded value would throw volatile-unconfigured.
		for (const ref of ["A1", "B1", "C1", "D1"]) {
			expect(await evaluateCell(book, "Results", ref)).toBe(4);
		}
	});

	it("does not invoke a UDF in a value excluded by the criteria", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "Data",
						rows: [
							["x", null, { formula: "PROBE()", value: 0 }],
							[null, null, 4],
						],
					},
					{
						name: "Results",
						rows: [
							[
								{
									formula: 'SUMIFS(Data!C1:C2,Data!A1:A2,"",Data!B1:B2,"")',
									value: 0,
								},
							],
						],
					},
				],
			}),
		);
		let visits = 0;
		expect(
			await evaluateCell(book, "Results", "A1", {
				functions: {
					PROBE: {
						minArgs: 0,
						maxArgs: 0,
						evaluate: () => {
							visits++;
							return 99;
						},
					},
				},
			}),
		).toBe(4);
		expect(visits).toBe(0);
	});
});
