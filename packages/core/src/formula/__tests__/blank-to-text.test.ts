import { describe, expect, it } from "vitest";
import { blankToText, errorValue, RangeView } from "../value";

const range = (width: number, value: string | number | null) =>
	new RangeView(
		"S",
		1,
		1,
		width,
		1,
		() => [],
		() => value,
	);

describe("blankToText", () => {
	it("reads an empty cell as an empty string", () => {
		expect(blankToText(null)).toBe("");
	});

	it("reads a range over one empty cell as an empty string", () => {
		expect(blankToText(range(1, null))).toBe("");
	});

	it("keeps a range over one non-empty cell, and any wider range, as is", () => {
		for (const view of [range(1, 5), range(1, ""), range(2, null), range(2, 5)]) {
			expect(blankToText(view)).toBe(view);
		}
	});

	it("keeps every other value as is", () => {
		for (const value of [0, -1, "", "x", true, false, errorValue("#N/A")]) {
			expect(blankToText(value)).toBe(value);
		}
	});
});
