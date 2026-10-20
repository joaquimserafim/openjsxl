import { describe, expect, it, vi } from "vitest";
import { MAX_ROW } from "../../ooxml/a1";
import { BUILTIN_ENTRIES } from "../builtins";
import type { EvalContext } from "../functions";
import { type EvalValue, errorValue, RangeView } from "../value";

const context: EvalContext = {
	date1904: false,
	now: () => 0,
	random: () => 0,
};

const mismatchedShapes: readonly [number, number][] = [
	[1, MAX_ROW - 1],
	[2, MAX_ROW],
];

const alignedCases: readonly [string, number][] = [
	["COUNTIFS", 2],
	["SUMIFS", 8],
	["AVERAGEIFS", 4],
];

function invoke(name: string, args: readonly EvalValue[]): EvalValue {
	const spec = BUILTIN_ENTRIES.find(([key]) => key === name)?.[1];
	if (spec === undefined || spec.lazyArgs) throw new Error(`missing eager builtin ${name}`);
	return spec.invoke(args, context);
}

// These views have real grid-sized extents but forbid access to any value. Shape rejection must
// happen through bounds alone, before even creating an entries iterator or resolving a criterion.
function unreadableRange(width: number, height: number) {
	const populated = vi.fn((): Iterable<readonly [number, number]> => {
		throw new Error("shape validation must not enumerate cells");
	});
	const resolve = vi.fn((): EvalValue => {
		throw new Error("shape validation must not resolve cells");
	});
	const range = new RangeView("S", 1, 1, width, height, populated, resolve);
	const readers = [
		vi.spyOn(range, "entries"),
		vi.spyOn(range, "values"),
		vi.spyOn(range, "cellAt"),
		vi.spyOn(range, "populatedCount"),
		vi.spyOn(range, "single"),
		vi.spyOn(range, "topLeft"),
		populated,
		resolve,
	];
	return { range, readers };
}

describe("conditional multi-range shape checks stay bounded", () => {
	it.each([
		"COUNTIFS",
		"SUMIFS",
		"AVERAGEIFS",
	])("%s rejects whole-column mismatches without reading cells", (name) => {
		for (const [width, height] of mismatchedShapes) {
			const expected = unreadableRange(1, MAX_ROW);
			const mismatch = unreadableRange(width, height);
			const args =
				name === "COUNTIFS"
					? [expected.range, ">0", mismatch.range, ">0"]
					: [expected.range, mismatch.range, ">0"];
			expect(invoke(name, args)).toEqual(errorValue("#VALUE!"));
			for (const reader of [...expected.readers, ...mismatch.readers]) {
				expect(reader).not.toHaveBeenCalled();
			}
		}
	});

	it.each([
		"COUNTIFS",
		"SUMIFS",
		"AVERAGEIFS",
	])("%s rejects malformed argument counts before touching range values", (name) => {
		const guarded = unreadableRange(1, MAX_ROW);
		const args =
			name === "COUNTIFS"
				? [guarded.range, ">0", guarded.range]
				: [guarded.range, guarded.range, ">0", guarded.range];
		expect(invoke(name, args)).toEqual(errorValue("#VALUE!"));
		expect(invoke(name, [])).toEqual(errorValue("#VALUE!"));
		for (const reader of guarded.readers) expect(reader).not.toHaveBeenCalled();
	});

	it.each(alignedCases)("%s still handles aligned sparse whole columns", (name, expected) => {
		const populated = vi.fn(
			(col: number): Iterable<readonly [number, number]> => [
				[col, 1],
				[col, MAX_ROW],
			],
		);
		const resolve = vi.fn((col: number, row: number): EvalValue => {
			if (row !== 1 && row !== MAX_ROW) {
				throw new Error("a sparse aggregate must not walk implicit blank positions");
			}
			return col === 3 ? (row === 1 ? 3 : 5) : 1;
		});
		const range = (col: number) =>
			new RangeView("S", col, 1, col, MAX_ROW, () => populated(col), resolve);
		const first = range(1);
		const second = range(2);
		const values = range(3);
		const args =
			name === "COUNTIFS" ? [first, ">0", second, ">0"] : [values, first, ">0", second, ">0"];
		expect(invoke(name, args)).toBe(expected);
		expect(populated.mock.calls.length).toBeLessThanOrEqual(6);
		expect(resolve.mock.calls.length).toBeLessThanOrEqual(12);
	});
});
