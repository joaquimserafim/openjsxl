import { describe, expect, it } from "vitest";
import {
	emptyCriterionRows,
	evaluateFormulas,
	expectCounts,
	mixedColumnRows,
	relationalRows,
	salesTableRows,
} from "./criteria-fixtures";

// F11.31. Oracles are Microsoft's function documentation:
// - COUNTIF: `=COUNTIF(A2:A5,"*")` "Counts the number of cells containing any text".
// - COUNTIFS: "If the criteria argument is a reference to an empty cell, the COUNTIFS function treats
//   the empty cell as a 0 value." (AVERAGEIF states the same.)
// Unsourced owner rulings: relational text criteria (<, >, <=, >=) match text cells only; an omitted
// criterion is 0; a formula's "" result is text. AVERAGEIFS's page also says an EMPTY CRITERIA-RANGE
// cell counts as 0; these pins keep blank criteria cells unmatched, as COUNTIF/COUNTIFS do.
// Every single-criterion count below is evaluated as both COUNTIF and COUNTIFS.

const GRID_ROWS = 1_048_576;

describe("wildcard criteria match text cells only", () => {
	it("counts a 100-row table like Excel, whole-column and bounded", async () => {
		await expectCounts(
			[
				["Data!A:A", '"*"', 91], // header + 90 names
				["Data!A2:A101", '"*"', 90],
				["Data!D:D", '"*"', 1], // only the "SKU" header is text
				["Data!D2:D101", '"*"', 0],
				["Data!A:A", '"<>*"', GRID_ROWS - 91],
				["Data!A2:A101", '"<>*"', 10],
			],
			salesTableRows(),
		);
		const values = await evaluateFormulas(
			[
				'COUNTIFS(Data!A:A,"*",Data!B:B,"<>South")',
				'COUNTIFS(Data!A2:A101,"*",Data!B2:B101,"<>South")',
				'SUMIFS(Data!C:C,Data!A:A,"*",Data!B:B,"<>South")',
				'COUNTIFS(Data!A:A,"*",Data!B:B,"North")',
				'SUMIF(Data!A:A,"*",Data!C:C)',
				'AVERAGEIF(Data!A2:A101,"*",Data!C2:C101)',
			],
			salesTableRows(),
		);
		expect(values).toEqual([61, 60, 2997, 30, 4500, 50]);
	});

	it('never matches numbers, booleans or blanks with "*", "?" or "2*", while "<>*" does', async () => {
		await expectCounts(
			[
				["Data!A1:A9", '"*"', 5], // apple, formula "", 2x, abcd, q
				["Data!A1:A9", '"?"', 1], // q, not the number 3
				["Data!A1:A9", '"2*"', 1], // 2x, not the number 2024
				["Data!A1:A9", '"????"', 1], // abcd, not 2024 or TRUE
				["Data!A1:A9", '"<>*"', 4], // 3, TRUE, blank, 2024
				["Data!A:A", '"*"', 5],
				["Data!A:A", '"<>*"', GRID_ROWS - 5],
				["Data!A:A", '"<>?"', GRID_ROWS - 1],
			],
			mixedColumnRows(),
		);
	});

	it('counts a formula returning "" as text for "*" (unsourced)', async () => {
		// No primary source settles this; an empty-string formula result is a string value.
		await expectCounts(
			[
				["Data!A1:A2", '"*"', 1],
				["Data!A1:A2", '"<>*"', 1],
			],
			[[{ formula: '""', value: "" }], [null]],
		);
	});
});

describe("relational text criteria match text cells only", () => {
	it("ignores numbers, booleans and blanks, whole-column and bounded", async () => {
		await expectCounts(
			[
				["Data!A:A", '"<z"', 1],
				["Data!A1:A3", '"<z"', 1],
				["Data!A:A", '">a"', 1],
				["Data!A1:A3", '">a"', 1],
				["Data!A:A", '"<=c"', 1],
				["Data!A:A", '">=c"', 1],
				["Data!A:A", '"<c"', 0],
				["Data!A1:A3", '">c"', 0],
			],
			relationalRows(),
		);
		const values = await evaluateFormulas(
			[
				'SUMIF(Data!A1:A3,"<z",Data!B1:B3)',
				'SUMIFS(Data!B:B,Data!A:A,"<z")',
				'COUNTIFS(Data!A:A,"<z",Data!B:B,">0")',
				'AVERAGEIFS(Data!B1:B3,Data!A1:A3,">=a")',
			],
			relationalRows(),
		);
		expect(values).toEqual([30, 30, 1, 30]);
	});
});

describe("an empty-cell or omitted criterion is the number 0", () => {
	const names = [{ name: "EmptyCell", refersTo: "Data!$E$1" }];

	it("treats every reference form to an empty cell alike", async () => {
		const values = await evaluateFormulas(
			[
				"COUNTIFS(Data!A1:A5,Data!E1,Data!B1:B5,Data!E1)",
				"COUNTIFS(Data!A1:A5,EmptyCell,Data!B1:B5,EmptyCell)",
				"COUNTIFS(Data!A1:A5,Data!E1:E1,Data!B1:B5,Data!E1:E1)",
				"COUNTIFS(Data!A1:A5,IF(TRUE,Data!E1),Data!B1:B5,IF(TRUE,Data!E1))",
				"COUNTIFS(Data!A:A,Data!E1,Data!B:B,Data!E1)",
				"SUMIFS(Data!C1:C5,Data!A1:A5,Data!E1,Data!B1:B5,Data!E1)",
				"AVERAGEIFS(Data!C1:C5,Data!A1:A5,Data!E1)",
				"SUMIF(Data!A1:A5,Data!E1,Data!C1:C5)",
			],
			emptyCriterionRows(),
			names,
		);
		expect(values).toEqual([2, 2, 2, 2, 2, 3, 1.5, 3]);
	});

	it("counts the same zeros through COUNTIF and COUNTIFS, including an omitted argument", async () => {
		await expectCounts(
			[
				["Data!A1:A5", "Data!E1", 2],
				["Data!A1:A5", "EmptyCell", 2],
				["Data!A1:A5", "Data!E1:E1", 2],
				["Data!A1:A5", "IF(TRUE,Data!E1)", 2],
				["Data!A:A", "Data!E1", 2],
				["Data!A1:A5", "", 2], // omitted: owner ruling (i), unsourced
			],
			emptyCriterionRows(),
			names,
		);
	});
});
