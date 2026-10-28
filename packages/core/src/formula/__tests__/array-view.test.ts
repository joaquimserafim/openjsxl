import { describe, expect, it } from "vitest";
import { scalarize } from "../coerce";
import {
	ArrayView,
	anchorValue,
	errorValue,
	gridPositions,
	mapArray,
	RangeView,
	sliceArray,
} from "../value";

describe("gridPositions", () => {
	it("walks a grid row by row from (1, 1)", () => {
		expect([...gridPositions(2, 2)]).toEqual([
			[1, 1],
			[2, 1],
			[1, 2],
			[2, 2],
		]);
		expect([...gridPositions(3, 1)]).toEqual([
			[1, 1],
			[2, 1],
			[3, 1],
		]);
	});

	it("yields nothing for an empty grid", () => {
		expect([...gridPositions(0, 3)]).toEqual([]);
		expect([...gridPositions(3, 0)]).toEqual([]);
	});
});

describe("ArrayView", () => {
	const view = new ArrayView([
		[1, "a"],
		[true, errorValue("#N/A")],
	]);

	it("spans its elements from (1, 1) and names no sheet", () => {
		expect([view.sheet, view.startCol, view.startRow, view.endCol, view.endRow]).toEqual([
			"",
			1,
			1,
			2,
			2,
		]);
		expect([view.width, view.height, view.cellCount]).toEqual([2, 2, 4]);
		expect(view).toBeInstanceOf(RangeView);
	});

	it("has every element populated, row by row", () => {
		expect(view.populatedCount()).toBe(4);
		expect([...view.entries()]).toEqual([
			{ col: 1, row: 1, value: 1 },
			{ col: 2, row: 1, value: "a" },
			{ col: 1, row: 2, value: true },
			{ col: 2, row: 2, value: errorValue("#N/A") },
		]);
		expect([...view.values()]).toEqual([1, "a", true, errorValue("#N/A")]);
	});

	it("reads elements by offset, and is empty outside them", () => {
		expect([view.cellAt(0, 0), view.cellAt(0, 1), view.cellAt(1, 0)]).toEqual([1, "a", true]);
		expect([view.cellAt(2, 0), view.cellAt(0, 2), view.cellAt(-1, 0)]).toEqual([
			null,
			null,
			null,
		]);
	});

	it("is one value only when it has one element", () => {
		expect(view.single()).toBeUndefined();
		expect(view.topLeft()).toBe(1);
		expect(new ArrayView([[7]]).single()).toBe(7);
	});
});

describe("anchorValue", () => {
	it("reads an array constant as its first element", () => {
		expect(anchorValue(new ArrayView([[3, 4]]))).toBe(3);
		expect(anchorValue(new ArrayView([["x"], ["y"]]))).toBe("x");
	});

	it("keeps any other value, a reference range included", () => {
		const range = new RangeView(
			"S",
			1,
			1,
			1,
			3,
			() => [],
			() => 7,
		);
		expect(anchorValue(range)).toBe(range);
		for (const value of [5, "", null, true, errorValue("#REF!")]) {
			expect(anchorValue(value)).toBe(value);
		}
	});
});

describe("scalarize", () => {
	it("keeps a scalar", () => {
		expect(scalarize(5)).toBe(5);
		expect(scalarize(null)).toBeNull();
	});

	it("reads a one-cell range as its cell", () => {
		expect(
			scalarize(
				new RangeView(
					"S",
					1,
					1,
					1,
					1,
					() => [],
					() => 7,
				),
			),
		).toBe(7);
	});

	it("reads an array constant of any size as its first element (F11.64)", () => {
		expect(scalarize(new ArrayView([[1, 2]]))).toBe(1);
		expect(scalarize(new ArrayView([[9]]))).toBe(9);
	});

	it("is #VALUE! for a reference range of more than one cell", () => {
		expect(
			scalarize(
				new RangeView(
					"S",
					1,
					1,
					1,
					3,
					() => [],
					() => 7,
				),
			),
		).toEqual(errorValue("#VALUE!"));
	});

	it("is #VALUE! for a one-cell range that resolves to itself", () => {
		const self: RangeView = new RangeView(
			"S",
			1,
			1,
			1,
			1,
			() => [],
			() => self,
		);
		expect(scalarize(self)).toEqual(errorValue("#VALUE!"));
	});
});

describe("mapArray", () => {
	it("keeps the shape and passes each element with its offsets", () => {
		const seen: string[] = [];
		const mapped = mapArray(
			new ArrayView([
				[1, 2],
				[3, 4],
			]),
			(v, row, col) => {
				seen.push(`${row},${col}`);
				return typeof v === "number" ? v * 10 : v;
			},
		);
		expect([mapped.width, mapped.height]).toEqual([2, 2]);
		expect([...mapped.values()]).toEqual([10, 20, 30, 40]);
		expect(seen).toEqual(["0,0", "0,1", "1,0", "1,1"]);
	});
});

describe("sliceArray", () => {
	const grid = new ArrayView([
		[1, 2, 3],
		[4, 5, 6],
	]);

	it("takes one row, one column, one element, or all of it", () => {
		const row = sliceArray(grid, 1, undefined);
		const col = sliceArray(grid, undefined, 2);
		const one = sliceArray(grid, 0, 1);
		const all = sliceArray(grid, undefined, undefined);
		expect([row.width, row.height, ...row.values()]).toEqual([3, 1, 4, 5, 6]);
		expect([col.width, col.height, ...col.values()]).toEqual([1, 2, 3, 6]);
		expect([one.width, one.height, ...one.values()]).toEqual([1, 1, 2]);
		expect([...all.values()]).toEqual([1, 2, 3, 4, 5, 6]);
	});
});
