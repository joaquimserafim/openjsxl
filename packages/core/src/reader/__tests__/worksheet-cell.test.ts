import { describe, expect, it } from "vitest";
import { effectiveStyle, rowDefaultStyleFromToken, safeColumn } from "../worksheet-cell";
import { createColumnStyleResolver } from "../worksheet-column-styles";

describe("worksheet cell helpers", () => {
	it("reads a valid column and tolerates a malformed ref", () => {
		expect(safeColumn("C99")).toBe(3);
		expect(safeColumn("aa2")).toBe(27);
		expect(safeColumn("bad")).toBeUndefined();
		expect(safeColumn("A0")).toBeUndefined();
		expect(safeColumn(`${"Z".repeat(64)}1`)).toBeUndefined();
	});

	it("uses a row style only when customFormat enables it", () => {
		expect(rowDefaultStyleFromToken({ s: "5", customFormat: "1" })).toBe(5);
		expect(rowDefaultStyleFromToken({ s: "6", customFormat: "true" })).toBe(6);
		expect(rowDefaultStyleFromToken({ s: "5" })).toBeUndefined();
		expect(rowDefaultStyleFromToken({ s: "bad", customFormat: "1" })).toBeUndefined();
	});

	it("resolves cell, row, and first matching column styles in order", () => {
		const columns = createColumnStyleResolver();
		columns.add({ min: 1, max: 2, style: 7 });
		columns.add({ min: 2, max: 3, style: 8 });
		expect(effectiveStyle("3", 5, 2, columns)).toBe(3);
		expect(effectiveStyle(undefined, 5, 2, columns)).toBe(5);
		expect(effectiveStyle(undefined, undefined, 2, columns)).toBe(7);
		expect(effectiveStyle(undefined, undefined, 4, columns)).toBeUndefined();

		// Declaring an invalid cell style still opts out of row/column inheritance.
		expect(effectiveStyle("bad", 5, 2, columns)).toBeUndefined();
	});
});
