import { describe, expect, it } from "vitest";
import { allStrings, legacyNumericStringValue } from "./numeric-literal-fixtures";

describe("numeric-literal fixtures", () => {
	it("keeps the 1.0.0 oracle's known answers", () => {
		expect([
			legacyNumericStringValue(" 1.5 "),
			legacyNumericStringValue("1..2"),
			legacyNumericStringValue("1e999"),
		]).toEqual([1.5, undefined, undefined]);
	});

	it("enumerates every string up to the length, shortest first", () => {
		expect([...allStrings("ab", 2)]).toEqual(["", "a", "b", "aa", "ab", "ba", "bb"]);
		expect([...allStrings("01.eE+-x ", 5)]).toHaveLength(66_430);
	});
});
