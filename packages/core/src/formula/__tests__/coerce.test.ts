import { describe, expect, it } from "vitest";
import { applyBinary, applyPostfix, applyUnary, compareValues, toNumber, toText } from "../coerce";
import { type EvalValue, errorValue, isErrorValue } from "../value";

// The Excel coercion matrix (M8 decision 6), pinned as unit tests. The end-to-end oracle check vs
// Python `formulas` lives in eval.test.ts; these lock the rules in isolation, including the traps:
// the empty-cell (null) vs empty-string ("") distinction, TRUE→1, no coercion in comparisons, and
// the number < text < FALSE < TRUE ordering.

const err = (code: Parameters<typeof errorValue>[0]) => errorValue(code);
const isErr = (v: EvalValue, code: string) => isErrorValue(v) && v.code === code;

describe("toNumber", () => {
	it("coerces per decision 6", () => {
		expect(toNumber(null)).toBe(0); // empty cell → 0
		expect(toNumber(true)).toBe(1);
		expect(toNumber(false)).toBe(0);
		expect(toNumber(5)).toBe(5);
		expect(toNumber("5")).toBe(5); // numeric string coerces
		expect(toNumber(" 5.5 ")).toBe(5.5);
		expect(toNumber("1e3")).toBe(1000);
	});

	it("rejects non-numeric and the empty STRING (distinct from the empty cell)", () => {
		expect(isErr(toNumber(""), "#VALUE!")).toBe(true); // empty string ≠ empty cell
		expect(isErr(toNumber("abc"), "#VALUE!")).toBe(true);
		expect(isErr(toNumber("1,000"), "#VALUE!")).toBe(true); // locale-invariant: no thousands sep
	});

	it("propagates an error value unchanged", () => {
		expect(toNumber(err("#N/A"))).toEqual(err("#N/A"));
	});
});

describe("toText", () => {
	it("coerces per decision 6", () => {
		expect(toText(null)).toBe(""); // empty cell → ""
		expect(toText(5)).toBe("5");
		expect(toText(5.5)).toBe("5.5");
		expect(toText(true)).toBe("TRUE");
		expect(toText(false)).toBe("FALSE");
		expect(toText("hi")).toBe("hi");
	});
});

describe("applyBinary — arithmetic", () => {
	it("computes and coerces operands", () => {
		expect(applyBinary("+", 10, 20)).toBe(30);
		expect(applyBinary("-", 10, 20)).toBe(-10);
		expect(applyBinary("*", 10, 20)).toBe(200);
		expect(applyBinary("/", 10, 20)).toBe(0.5);
		expect(applyBinary("^", 10, 2)).toBe(100);
		expect(applyBinary("+", "5", true)).toBe(6); // "5"→5, TRUE→1
		expect(applyBinary("+", null, 1)).toBe(1); // empty→0
	});

	it("returns #DIV/0! for division (and 0^negative) by zero", () => {
		expect(isErr(applyBinary("/", 1, 0), "#DIV/0!")).toBe(true);
		expect(isErr(applyBinary("^", 0, -1), "#DIV/0!")).toBe(true);
	});

	it("returns #NUM! on overflow / non-real", () => {
		expect(isErr(applyBinary("^", 1e308, 10), "#NUM!")).toBe(true);
		expect(isErr(applyBinary("^", -2, 0.5), "#NUM!")).toBe(true);
	});

	it("propagates the leftmost error operand", () => {
		expect(applyBinary("+", err("#DIV/0!"), err("#N/A"))).toEqual(err("#DIV/0!"));
	});
});

describe("applyBinary — concatenation", () => {
	it("joins text with empty→'' and General number text", () => {
		expect(applyBinary("&", 10, 20)).toBe("1020");
		expect(applyBinary("&", "a", null)).toBe("a");
		expect(applyBinary("&", true, "!")).toBe("TRUE!");
	});
});

describe("applyBinary — comparisons never coerce across types", () => {
	it("orders number < text < FALSE < TRUE", () => {
		expect(applyBinary("<", 1, "1")).toBe(true); // number < text
		expect(applyBinary("=", 1, "1")).toBe(false); // different types are not equal
		expect(applyBinary("<", "z", true)).toBe(true); // text < boolean
		expect(applyBinary("<", false, true)).toBe(true);
	});

	it("treats an empty cell as the other operand's type", () => {
		expect(applyBinary("=", null, 0)).toBe(true); // empty = 0
		expect(applyBinary("=", null, "")).toBe(true); // empty = ""
		expect(applyBinary("=", null, false)).toBe(true); // empty = FALSE
		expect(applyBinary("=", null, null)).toBe(true);
	});

	it("compares text case-insensitively", () => {
		expect(applyBinary("=", "ABC", "abc")).toBe(true);
		expect(applyBinary("<", "a", "b")).toBe(true);
	});

	it("orders numbers and reports every relational operator", () => {
		expect(applyBinary(">", 20, 10)).toBe(true);
		expect(applyBinary("<>", 1, 2)).toBe(true);
		expect(applyBinary(">=", 5, 5)).toBe(true);
		expect(applyBinary("<=", 5, 6)).toBe(true);
	});
});

describe("compareValues", () => {
	it("returns -1/0/1 and propagates errors", () => {
		expect(compareValues(1, 2)).toBe(-1);
		expect(compareValues(2, 2)).toBe(0);
		expect(compareValues(3, 2)).toBe(1);
		expect(compareValues(err("#REF!"), 1)).toEqual(err("#REF!"));
	});
});

describe("applyUnary / applyPostfix", () => {
	it("negates, coerces with unary plus, and reduces @ to a scalar", () => {
		expect(applyUnary("-", 5)).toBe(-5);
		expect(applyUnary("-", "5")).toBe(-5);
		expect(applyUnary("+", true)).toBe(1);
		expect(isErr(applyUnary("-", "x"), "#VALUE!")).toBe(true);
		expect(applyUnary("@", 5)).toBe(5); // scalar passthrough
	});

	it("applies percent and degrades spill to #REF!", () => {
		expect(applyPostfix("%", 50)).toBe(0.5);
		expect(isErr(applyPostfix("#", 5), "#REF!")).toBe(true); // dynamic-array spill not evaluated
	});
});
