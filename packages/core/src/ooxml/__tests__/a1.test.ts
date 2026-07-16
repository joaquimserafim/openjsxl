import { describe, expect, it } from "vitest";
import {
	columnToIndex,
	formatRef,
	indexToColumn,
	parseCanonicalCell,
	parseCanonicalRange,
	parseRef,
} from "../a1";

describe("columnToIndex", () => {
	it("maps the bijective base-26 boundaries", () => {
		expect(columnToIndex("A")).toBe(1);
		expect(columnToIndex("Z")).toBe(26);
		expect(columnToIndex("AA")).toBe(27);
		expect(columnToIndex("AZ")).toBe(52);
		expect(columnToIndex("BA")).toBe(53);
		expect(columnToIndex("XFD")).toBe(16384); // Excel's last column
	});

	it("is case-insensitive", () => {
		expect(columnToIndex("aa")).toBe(27);
	});

	it("rejects invalid input", () => {
		expect(() => columnToIndex("")).toThrow();
		expect(() => columnToIndex("A1")).toThrow();
		expect(() => columnToIndex("-")).toThrow();
	});

	it("rejects an overflowing ref instead of returning a non-integer", () => {
		// A ref far past XFD used to overflow silently to a lossy float / Infinity, poisoning
		// downstream column arithmetic. It must throw so callers can reject or fall back.
		expect(() => columnToIndex("A".repeat(300))).toThrow();
		// 13 'A's is the first length that crosses MAX_SAFE_INTEGER in bijective base-26.
		expect(() => columnToIndex("A".repeat(13))).toThrow();
		// A ref that still maps within safe-integer range is accepted (returns a finite integer).
		expect(Number.isSafeInteger(columnToIndex("A".repeat(9)))).toBe(true);
	});
});

describe("indexToColumn", () => {
	it("round-trips every index from 1 to 20000", () => {
		for (let i = 1; i <= 20000; i++) {
			expect(columnToIndex(indexToColumn(i))).toBe(i);
		}
	});

	it("rejects non-positive indices", () => {
		expect(() => indexToColumn(0)).toThrow();
	});
});

describe("parseRef / formatRef", () => {
	it("parses and formats A1 references", () => {
		expect(parseRef("A1")).toEqual({ col: 1, row: 1 });
		expect(parseRef("XFD1048576")).toEqual({ col: 16384, row: 1048576 });
		expect(formatRef({ col: 1, row: 1 })).toBe("A1");
	});

	it("round-trips references", () => {
		for (const ref of ["A1", "B2", "AA10", "XFD1048576"]) {
			expect(formatRef(parseRef(ref))).toBe(ref);
		}
	});
});

describe("parseCanonicalCell", () => {
	it("accepts a canonical, in-grid single cell", () => {
		expect(parseCanonicalCell("A1")).toEqual({ col: 1, row: 1 });
		expect(parseCanonicalCell("XFD1048576")).toEqual({ col: 16384, row: 1048576 });
	});

	it("rejects lowercase, leading zeros, out-of-grid, and non-A1 tokens", () => {
		expect(parseCanonicalCell("a1")).toBeUndefined();
		expect(parseCanonicalCell("A01")).toBeUndefined();
		expect(parseCanonicalCell("XFE1")).toBeUndefined(); // past last column
		expect(parseCanonicalCell("A1048577")).toBeUndefined(); // past last row
		expect(parseCanonicalCell("A1:B2")).toBeUndefined(); // a range, not a cell
		expect(parseCanonicalCell("")).toBeUndefined();
	});
});

describe("parseCanonicalRange", () => {
	it("parses a canonical range to its two corners (as written, not sorted)", () => {
		expect(parseCanonicalRange("A1:C3")).toEqual({
			from: { col: 1, row: 1 },
			to: { col: 3, row: 3 },
		});
	});

	it("treats a single cell as a degenerate range", () => {
		expect(parseCanonicalRange("B2")).toEqual({
			from: { col: 2, row: 2 },
			to: { col: 2, row: 2 },
		});
	});

	it("rejects when either end is non-canonical or out of grid", () => {
		expect(parseCanonicalRange("a1:c3")).toBeUndefined();
		expect(parseCanonicalRange("A1:ZZZZ9")).toBeUndefined();
		expect(parseCanonicalRange("A1:B1048577")).toBeUndefined();
		expect(parseCanonicalRange("A1:B2:C3")).toBeUndefined(); // malformed
	});
});
