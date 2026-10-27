import { describe, expect, it } from "vitest";
import { blankToZero, errorValue, RangeView } from "../value";

describe("blankToZero", () => {
	it("reads an empty cell as 0", () => {
		expect(blankToZero(null)).toBe(0);
	});

	it("keeps every other value as is", () => {
		for (const value of [0, -1, 42, "", "x", true, false, errorValue("#N/A")]) {
			expect(blankToZero(value)).toBe(value);
		}
	});

	it("keeps a range a range, even over one empty cell", () => {
		const range = new RangeView(
			"S",
			1,
			1,
			1,
			1,
			() => [],
			() => null,
		);
		expect(blankToZero(range)).toBe(range);
	});
});
