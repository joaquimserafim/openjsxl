import { describe, expect, it } from "vitest";
import { isNonCellContent } from "../content";

describe("isNonCellContent", () => {
	it("names comments and anchored shapes, not the cell's own text elements", () => {
		expect(
			["office:annotation", "draw:frame", "draw:custom-shape", "draw:text-box"].map(
				isNonCellContent,
			),
		).toEqual([true, true, true, true]);
		expect(
			["text:p", "text:a", "text:s", "office:annotation-end"].map(isNonCellContent),
		).toEqual([false, false, false, false]);
	});
});
