import { describe, expect, it } from "vitest";
import { indexArrayView, replaceArrayErrors } from "../builtins";
import { ArrayView, errorValue, isRangeView } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.64, from its review: the built-ins that treat an array constant differently from a range.

const values = (v: unknown) => (isRangeView(v) ? [v.width, v.height, ...v.values()] : v);
const grid = new ArrayView([
	[1, 2],
	[3, 4],
]);

describe("replaceArrayErrors", () => {
	it("replaces each caught error element with the fallback, evaluated once", () => {
		let calls = 0;
		const out = replaceArrayErrors(
			new ArrayView([[errorValue("#N/A"), 1, errorValue("#DIV/0!")]]),
			() => true,
			() => {
				calls++;
				return 0;
			},
		);
		expect(values(out)).toEqual([3, 1, 0, 1, 0]);
		expect(calls).toBe(1);
	});

	it("keeps errors the caller does not catch", () => {
		const out = replaceArrayErrors(
			new ArrayView([[errorValue("#N/A"), errorValue("#DIV/0!")]]),
			(e) => e.code === "#N/A",
			() => "x",
		);
		expect(values(out)).toEqual([2, 1, "x", errorValue("#DIV/0!")]);
	});

	it("never runs the fallback when nothing is caught, and reads an empty fallback as an empty string", () => {
		let calls = 0;
		const clean = replaceArrayErrors(
			new ArrayView([[1, 2]]),
			() => true,
			() => {
				calls++;
				return 0;
			},
		);
		expect([values(clean), calls]).toEqual([[2, 1, 1, 2], 0]);
		expect(
			values(
				replaceArrayErrors(
					new ArrayView([[errorValue("#N/A")]]),
					() => true,
					() => null,
				),
			),
		).toEqual([1, 1, ""]);
		expect(
			values(
				replaceArrayErrors(new ArrayView([[errorValue("#N/A")]]), () => true, undefined),
			),
		).toEqual([1, 1, ""]);
	});
});

describe("indexArrayView", () => {
	it("returns a column, a row, or the whole array for a 0 index", () => {
		expect(values(indexArrayView(grid, 0, 1))).toEqual([1, 2, 1, 3]);
		expect(values(indexArrayView(grid, 2, 0))).toEqual([2, 1, 3, 4]);
		expect(values(indexArrayView(grid, 0, 0))).toEqual([2, 2, 1, 2, 3, 4]);
		const row = new ArrayView([[1, 2, 3]]);
		expect(indexArrayView(row, 0, undefined)).toBe(row);
	});

	it("returns a whole row for one index on a 2-D array", () => {
		expect(values(indexArrayView(grid, 2, undefined))).toEqual([2, 1, 3, 4]);
	});

	it("is #REF! past the array or below 0", () => {
		for (const [row, col] of [
			[3, 0],
			[0, 3],
			[3, undefined],
			[-1, 0],
			[0, -1],
		] as const) {
			expect(indexArrayView(grid, row, col)).toEqual(errorValue("#REF!"));
		}
	});

	it("leaves one element, and a 1-D index, to INDEX's range path", () => {
		expect(indexArrayView(grid, 2, 1)).toBeUndefined();
		expect(indexArrayView(new ArrayView([[1, 2, 3]]), 2, undefined)).toBeUndefined();
	});
});

describe("built-ins over an array constant", () => {
	it("IFERROR and IFNA replace error elements one by one", async () => {
		expect(
			await evaluateFormulas(
				[
					"IFERROR({#N/A,1},0)",
					'IFNA({#N/A,1},"x")',
					"IFERROR(IF(TRUE,{#N/A,1}),0)",
					"SUM(IFERROR({1,#N/A},0))",
					"SUM(IFNA({1,#N/A},0))",
					"AVERAGE(IFERROR({2,#N/A},0))",
					"IFNA({#DIV/0!,1},0)",
				],
				[],
			),
		).toEqual([0, "x", 0, 1, 1, 1, errorValue("#DIV/0!")]);
	});

	it("IFERROR still evaluates the fallback only when an element is an error", async () => {
		expect(await evaluateFormulas(["IFERROR({1,2},NOW())"], [])).toEqual([1]);
		await expect(evaluateFormulas(["IFERROR({1,#N/A},NOW())"], [])).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
	});

	it("an error that is not first still reaches an aggregate", async () => {
		expect(await evaluateFormulas(["SUM({1,#N/A})", "MAX({2,#DIV/0!})"], [])).toEqual([
			errorValue("#N/A"),
			errorValue("#DIV/0!"),
		]);
	});

	it("MATCH gives #N/A for a 2-D constant and looks along a 1-D one", async () => {
		expect(
			await evaluateFormulas(
				[
					"MATCH(3,{1,2;3,4},0)",
					"MATCH(2,{1,2;3,4},0)",
					"MATCH(2,{1;2;3},0)",
					"MATCH(3,{1,2,3},0)",
				],
				[],
			),
		).toEqual([errorValue("#N/A"), errorValue("#N/A"), 2, 3]);
	});

	it("CONCATENATE reads its first element, where CONCAT joins every element", async () => {
		expect(
			await evaluateFormulas(
				['CONCATENATE({"a","b"})', 'CONCATENATE({"a","b"},"c")', 'CONCAT({"a","b"})'],
				[],
			),
		).toEqual(["a", "ac", "ab"]);
	});

	it("INDEX returns a whole row or column for 0 or one index on a 2-D constant", async () => {
		expect(
			await evaluateFormulas(
				[
					"INDEX({1,2;3,4},0,1)",
					"SUM(INDEX({1,2;3,4},0,1))",
					"SUM(INDEX({1,2;3,4},2))",
					"SUM(INDEX({1,2;3,4},0,0))",
					"SUM(INDEX({1,2,3},0))",
					"INDEX({1,2;3,4},3,0)",
				],
				[],
			),
		).toEqual([1, 4, 7, 10, 6, errorValue("#REF!")]);
	});
});
