import { describe, expect, it } from "vitest";
import { retainedHeapBytes } from "../../ooxml/__tests__/retained-heap";
import { collapseQuotes } from "../parse";

describe("collapseQuotes", () => {
	it("collapses each escaped quote in the range to one quote", () => {
		expect(collapseQuotes('x"a""b""c"', 2, 9)).toBe('a"b"c');
		expect(collapseQuotes('""""', 0, 4)).toBe('""');
	});

	it("returns the range itself when it holds no escaped quote, and stops at its end", () => {
		expect(collapseQuotes("abc", 0, 3)).toBe("abc");
		expect(collapseQuotes('ab"",c""d', 0, 4)).toBe('ab"');
		expect(collapseQuotes("abc", 1, 1)).toBe("");
	});

	it("matches replaceAll across more pairs than one join batch", () => {
		const body = 'ab""c'.repeat(10_000);
		expect(collapseQuotes(body, 0, body.length)).toBe(body.replaceAll('""', '"'));
	});

	it("builds a flat string, about its own size in memory", () => {
		// replaceAll builds a rope of about two nodes per pair: near 28 bytes per character.
		const body = Array.from({ length: 500_000 }, () => 'ab""').join("");
		const retained = retainedHeapBytes(() => collapseQuotes(body, 0, body.length));
		expect(retained).toBeLessThan(2 * body.length);
	});
});
