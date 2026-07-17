import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import {
	columnToIndex,
	formatRef,
	indexToColumn,
	parseCanonicalCell,
	parseCanonicalRange,
	parseRef,
} from "../a1";

// These four helpers are exported from the package index, so their thrown type freezes at 1.0:
// every rejection must be the typed XlsxError("invalid-input") the public error contract promises.
// A bare `toThrow()` would keep passing if one regressed to `new Error`, so assert the CODE. Sync
// analogue of the writer suite's async `writeErr` (writer/__tests__/protection.test.ts).
function refErr(fn: () => unknown): XlsxError {
	try {
		fn();
	} catch (e) {
		if (e instanceof XlsxError) return e;
		throw e;
	}
	throw new Error("expected a throw");
}

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

	it("rejects invalid input with a typed XlsxError", () => {
		expect(refErr(() => columnToIndex("")).code).toBe("invalid-input");
		expect(refErr(() => columnToIndex("A1")).code).toBe("invalid-input");
		expect(refErr(() => columnToIndex("-")).code).toBe("invalid-input");
	});

	it("rejects an overflowing ref instead of returning a non-integer", () => {
		// A ref far past XFD used to overflow silently to a lossy float / Infinity, poisoning
		// downstream column arithmetic. It must throw so callers can reject or fall back.
		expect(refErr(() => columnToIndex("A".repeat(300))).code).toBe("invalid-input");
		// 13 'A's is the first length that crosses MAX_SAFE_INTEGER in bijective base-26.
		expect(refErr(() => columnToIndex("A".repeat(13))).code).toBe("invalid-input");
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

	it("rejects non-positive indices with a typed XlsxError", () => {
		expect(refErr(() => indexToColumn(0)).code).toBe("invalid-input");
		expect(refErr(() => indexToColumn(-1)).code).toBe("invalid-input");
		expect(refErr(() => indexToColumn(1.5)).code).toBe("invalid-input");
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

	it("rejects a non-A1 token / bad row with a typed XlsxError", () => {
		expect(refErr(() => parseRef("ZZ")).code).toBe("invalid-input"); // no row
		expect(refErr(() => parseRef("A0")).code).toBe("invalid-input"); // row 0 is not A1
		expect(refErr(() => parseRef("1A")).code).toBe("invalid-input"); // reversed
		expect(refErr(() => formatRef({ col: 1, row: 0 })).code).toBe("invalid-input");
		expect(refErr(() => formatRef({ col: 1, row: 1.5 })).code).toBe("invalid-input");
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
