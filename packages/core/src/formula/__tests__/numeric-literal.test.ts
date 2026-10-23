import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { numericStringValue } from "../coerce";
import { allStrings, legacyNumericStringValue } from "./numeric-literal-fixtures";

// F11.51: the en-US numeric-literal check is linear in text length and accepts the same literals as
// 1.0.0. toNumber, VALUE, COUNT and criteria parsing all go through numericStringValue.

const MATRIX: readonly (readonly [string, number | undefined])[] = [
	["1", 1],
	["1.", 1],
	["1.5", 1.5],
	[".5", 0.5],
	["1e5", 100_000],
	["1.e5", 100_000],
	["-1.5E-3", -0.0015],
	["+.5", 0.5],
	["1..2", undefined],
	[".", undefined],
	["", undefined],
	["1e", undefined],
	["e5", undefined],
	[" 1.5 ", 1.5],
	["1e999", undefined],
];

describe("numericStringValue", () => {
	it.each(MATRIX)("reads %j as %s, as 1.0.0 did", (text, expected) => {
		expect(numericStringValue(text)).toBe(expected);
		expect(legacyNumericStringValue(text)).toBe(expected);
	});

	it("agrees with 1.0.0 on every string of up to 5 literal characters", () => {
		const mismatches: string[] = [];
		for (const text of allStrings("01.eE+-x ", 5)) {
			if (!Object.is(numericStringValue(text), legacyNumericStringValue(text)))
				mismatches.push(text);
		}
		expect(mismatches).toEqual([]);
	});

	// Both shapes are quadratic in 1.0.0 (ratio about 4).
	it.each([
		["a digit run then a letter", (n: number) => `${"1".repeat(n)}x`],
		["a digit run then a point and a letter", (n: number) => `${"1".repeat(n)}.x`],
	])(
		"is linear in text length on %s",
		(_name, shape) => {
			const ratio = bestDoublingRatio((n) => () => numericStringValue(shape(n)), 25_000);
			expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
		},
		60_000,
	);
});
