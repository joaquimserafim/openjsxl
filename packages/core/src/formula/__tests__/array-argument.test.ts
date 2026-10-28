import { describe, expect, it } from "vitest";
import type { ArrayElement } from "../ast";
import { arrayArgument, arrayElementValue } from "../eval";
import { parseFormula } from "../parser";
import { ArrayView, errorValue } from "../value";

const num = (value: number): ArrayElement => ({ type: "number", value, raw: String(value) });

describe("arrayElementValue", () => {
	it("reads a literal element", () => {
		expect(arrayElementValue(num(2))).toBe(2);
		expect(arrayElementValue({ type: "string", value: "a" })).toBe("a");
		expect(arrayElementValue({ type: "boolean", value: false })).toBe(false);
		expect(arrayElementValue({ type: "error", value: "#N/A" })).toEqual(errorValue("#N/A"));
	});

	it("applies a sign to a number, and to a numeric string", () => {
		expect(arrayElementValue({ type: "unary", op: "-", operand: num(1) })).toBe(-1);
		expect(arrayElementValue({ type: "unary", op: "+", operand: num(3) })).toBe(3);
		expect(
			arrayElementValue({ type: "unary", op: "-", operand: { type: "string", value: "2" } }),
		).toBe(-2);
	});

	it("is an error for a sign on text, and keeps an error under a sign", () => {
		expect(
			arrayElementValue({ type: "unary", op: "-", operand: { type: "string", value: "a" } }),
		).toEqual(errorValue("#VALUE!"));
		expect(
			arrayElementValue({
				type: "unary",
				op: "-",
				operand: { type: "error", value: "#DIV/0!" },
			}),
		).toEqual(errorValue("#DIV/0!"));
	});

	it("is #VALUE! for a sign on anything but a literal", () => {
		expect(
			arrayElementValue({
				type: "unary",
				op: "-",
				operand: { type: "cell", ref: "A1" },
			} as never),
		).toEqual(errorValue("#VALUE!"));
	});
});

describe("arrayArgument", () => {
	it("turns an array constant into a view of its elements", () => {
		const view = arrayArgument(parseFormula('{1,-2;TRUE,"x"}'));
		expect(view).toBeInstanceOf(ArrayView);
		expect(view === undefined ? [] : [...view.values()]).toEqual([1, -2, true, "x"]);
		expect([view?.width, view?.height]).toEqual([2, 2]);
	});

	it("looks through parentheses", () => {
		const view = arrayArgument(parseFormula("(({1,2}))"));
		expect(view === undefined ? [] : [...view.values()]).toEqual([1, 2]);
	});

	it("builds the view of a node once", () => {
		const node = parseFormula("{1,2,3}");
		const first = arrayArgument(node);
		expect(first).toBeInstanceOf(ArrayView);
		expect(arrayArgument(node)).toBe(first);
		expect(arrayArgument(parseFormula("{1,2,3}"))).not.toBe(first);
	});

	it("leaves every other argument alone", () => {
		for (const formula of ["A1:A3", "1", '"a"', "{1,2}*2", "-{1,2}", "SUM({1,2})", "(A1)"]) {
			expect(arrayArgument(parseFormula(formula))).toBeUndefined();
		}
	});
});
