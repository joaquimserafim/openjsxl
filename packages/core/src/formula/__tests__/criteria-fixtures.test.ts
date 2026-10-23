import { describe, expect, it } from "vitest";
import {
	countForms,
	emptyCriterionRows,
	evaluateFormulas,
	expectCounts,
	mixedColumnRows,
	relationalRows,
	salesTableRows,
} from "./criteria-fixtures";

describe("criteria test fixtures", () => {
	it("evaluates formulas against the Data sheet, with optional defined names", async () => {
		const values = await evaluateFormulas(
			["SUM(Data!A1:A2)", "SUM(Pair)*2"],
			[[1], [2]],
			[{ name: "Pair", refersTo: "Data!$A$1:$A$2" }],
		);
		expect(values).toEqual([3, 6]);
	});

	it("writes one criterion as COUNTIF and COUNTIFS", () => {
		expect(countForms("Data!A:A", '"*"')).toEqual([
			'COUNTIF(Data!A:A,"*")',
			'COUNTIFS(Data!A:A,"*")',
		]);
	});

	it("asserts both count forms against each expected count", async () => {
		await expectCounts([["Data!A1:A2", '">0"', 2]], [[1], [2]]);
		await expect(expectCounts([["Data!A1:A2", '">0"', 3]], [[1], [2]])).rejects.toThrow();
	});

	it("builds a header plus 100 rows with 10 blank names", () => {
		const rows = salesTableRows();
		expect(rows).toHaveLength(101);
		expect(rows[0]).toEqual(["Name", "Region", "Amount", "SKU"]);
		expect(rows.filter((row) => row[0] === null)).toHaveLength(10);
		expect(rows[3]).toEqual(["N3", "South", 3, 1003]);
		expect(rows[10]).toEqual([null, "North", 10, 1010]);
	});

	it("builds the mixed, relational and empty-criterion columns", () => {
		expect(mixedColumnRows()).toHaveLength(9);
		expect(relationalRows().map((row) => row[0])).toEqual([5, true, "c"]);
		expect(emptyCriterionRows().map((row) => row[0])).toEqual([0, 0, null, 5, "t"]);
	});
});
