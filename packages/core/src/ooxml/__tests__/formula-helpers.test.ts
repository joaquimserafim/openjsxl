import { describe, expect, it } from "vitest";
import { MAX_COL, MAX_ROW } from "../a1";
import {
	bracketSpanEnd,
	refCanStartAfter,
	shiftColRange,
	shiftRef,
	shiftRowRange,
} from "../formula";

// Direct unit tests for each translator helper (F11.9 follow-up, owner request): every branch of
// every helper is pinned on its own, so a regression is localized to the function that broke rather
// than surfacing only as a translateFormula vector mismatch. These are internal seams — exported
// for tests only, never part of the public API (the export-surface pin guards that).

describe("refCanStartAfter — a reference may only start at a formula boundary", () => {
	const code = (s: string) => s.charCodeAt(0);

	it("allows the start of the text and every operator/delimiter", () => {
		expect(refCanStartAfter(-1)).toBe(true);
		for (const ch of [
			"+",
			"-",
			"*",
			"/",
			"^",
			"&",
			"=",
			"<",
			">",
			"(",
			")",
			",",
			";",
			":",
			"!",
			" ",
			"\t",
			"\n",
			'"',
			"'",
			"%",
			"{",
			"}",
			"]",
			"@",
			"#",
			"$",
		]) {
			expect(refCanStartAfter(code(ch)), JSON.stringify(ch)).toBe(true);
		}
	});

	it("forbids identifier-ish predecessors: digits, ASCII letters, `_`, `.`, and any non-ASCII unit", () => {
		for (const ch of ["0", "5", "9", "A", "Z", "a", "z", "_", "."]) {
			expect(refCanStartAfter(code(ch)), JSON.stringify(ch)).toBe(false);
		}
		expect(refCanStartAfter(0x80)).toBe(false); // first non-ASCII code unit
		expect(refCanStartAfter(code("ñ"))).toBe(false);
		expect(refCanStartAfter(code("日"))).toBe(false);
		expect(refCanStartAfter(0xd83d)).toBe(false); // a surrogate half
	});
});

describe("shiftRef — one cell reference, honoring $ pins", () => {
	it("shifts relative parts, pins absolute parts, and canonicalizes to uppercase", () => {
		expect(shiftRef("A1", 2, 1)).toBe("B3");
		expect(shiftRef("a1", 2, 1)).toBe("B3");
		expect(shiftRef("$A$1", 2, 1)).toBe("$A$1");
		expect(shiftRef("$a$1", 2, 1)).toBe("$A$1"); // fully pinned still canonicalized
		expect(shiftRef("$A1", 3, 2)).toBe("$A4");
		expect(shiftRef("A$1", 3, 2)).toBe("C$1");
		expect(shiftRef("AA10", 2, 2)).toBe("AC12");
		expect(shiftRef("B2", -1, -1)).toBe("A1");
	});

	it("is the identity for a zero delta", () => {
		expect(shiftRef("C3", 0, 0)).toBe("C3");
		expect(shiftRef("c3", 0, 0)).toBe("C3"); // …up to canonical case
	});

	it("returns undefined for text that is not an addressable cell", () => {
		expect(shiftRef("ZZZ1", 1, 0)).toBeUndefined(); // pseudo-column past XFD
		expect(shiftRef(`A${MAX_ROW + 1}`, 0, 0)).toBeUndefined(); // row past the grid
		expect(shiftRef("A0", 1, 0)).toBeUndefined(); // rows are 1-based
		expect(shiftRef("A01", 1, 0)).toBeUndefined(); // leading zero: a name, not a cell
		expect(shiftRef("name", 1, 0)).toBeUndefined();
		expect(shiftRef("A1:B2", 1, 0)).toBeUndefined(); // a range, not a single ref
		expect(shiftRef("", 1, 0)).toBeUndefined();
	});

	it("rewrites a shift that lands off the grid to #REF!", () => {
		expect(shiftRef("A1", -1, 0)).toBe("#REF!"); // row 0
		expect(shiftRef("A1", 0, -1)).toBe("#REF!"); // column 0
		expect(shiftRef(`A${MAX_ROW}`, 1, 0)).toBe("#REF!"); // past the last row
		expect(shiftRef("XFD1", 0, 1)).toBe("#REF!"); // past the last column
		expect(shiftRef("$A$1", -5, -5)).toBe("$A$1"); // pinned parts never move, so never off-grid
	});

	it("addresses the grid corners exactly", () => {
		expect(shiftRef("A1", MAX_ROW - 1, MAX_COL - 1)).toBe(`XFD${MAX_ROW}`);
		expect(shiftRef(`XFD${MAX_ROW}`, -(MAX_ROW - 1), -(MAX_COL - 1))).toBe("A1");
	});
});

describe("shiftColRange — a whole-column range", () => {
	it("shifts each relative endpoint and pins absolute ones, uppercase out", () => {
		expect(shiftColRange("A:A", 1)).toBe("B:B");
		expect(shiftColRange("a:c", 1)).toBe("B:D");
		expect(shiftColRange("$A:$B", 2)).toBe("$A:$B");
		expect(shiftColRange("A:$B", 2)).toBe("C:$B");
		expect(shiftColRange("$A:B", 2)).toBe("$A:D");
	});

	it("leaves a non-range or an endpoint past XFD verbatim", () => {
		expect(shiftColRange("A1", 1)).toBe("A1");
		expect(shiftColRange("ZZZ:A", 1)).toBe("ZZZ:A");
		expect(shiftColRange("A:ZZZ", 1)).toBe("A:ZZZ");
	});

	it("rewrites an off-grid shift to #REF!", () => {
		expect(shiftColRange("A:B", -1)).toBe("#REF!");
		expect(shiftColRange("XFC:XFD", 1)).toBe("#REF!");
		expect(shiftColRange("XFC:XFD", 0)).toBe("XFC:XFD");
	});
});

describe("shiftRowRange — a whole-row range", () => {
	it("shifts each relative endpoint and pins absolute ones", () => {
		expect(shiftRowRange("1:1", 1)).toBe("2:2");
		expect(shiftRowRange("2:5", 2)).toBe("4:7");
		expect(shiftRowRange("$2:$5", 3)).toBe("$2:$5");
		expect(shiftRowRange("2:$5", 3)).toBe("5:$5");
	});

	it("leaves a non-range, row 0, or an endpoint past the grid verbatim", () => {
		expect(shiftRowRange("A1", 1)).toBe("A1");
		expect(shiftRowRange("0:0", 1)).toBe("0:0");
		expect(shiftRowRange(`1:${MAX_ROW + 1}`, 1)).toBe(`1:${MAX_ROW + 1}`);
	});

	it("rewrites an off-grid shift to #REF!", () => {
		expect(shiftRowRange("1:2", -1)).toBe("#REF!");
		expect(shiftRowRange(`${MAX_ROW - 1}:${MAX_ROW}`, 1)).toBe("#REF!");
	});
});

describe("bracketSpanEnd — the balanced span from an opening `[`", () => {
	it("returns the exclusive end of a simple, nested, and consecutive span", () => {
		expect(bracketSpanEnd("[Col]", 0)).toBe(5);
		expect(bracketSpanEnd("T[Col]+A1", 1)).toBe(6);
		expect(bracketSpanEnd("T[[#Data],[A1]]", 1)).toBe(15);
		expect(bracketSpanEnd("[a][b]", 0)).toBe(3); // only the first span
		expect(bracketSpanEnd("[a][b]", 3)).toBe(6);
	});

	it("treats a `'` as escaping the next character, whatever it is", () => {
		expect(bracketSpanEnd("[Col']]", 0)).toBe(7); // `']` is the escaped `]`; the real close follows
		expect(bracketSpanEnd("[Col'[1']]", 0)).toBe(10);
		expect(bracketSpanEnd("[a'']", 0)).toBe(5); // `''` is an escaped quote
		expect(bracketSpanEnd("[a']", 0)).toBe(-1); // the escape eats the only `]` — unclosed
		expect(bracketSpanEnd("[a'", 0)).toBe(-1); // escape at the very end
	});

	it("returns -1 for an unbalanced span", () => {
		expect(bracketSpanEnd("[", 0)).toBe(-1);
		expect(bracketSpanEnd("[[a]", 0)).toBe(-1);
		expect(bracketSpanEnd("T[Col", 1)).toBe(-1);
		expect(bracketSpanEnd("", 0)).toBe(-1);
	});

	it("is linear in the span length", () => {
		const big = `[${"[x],".repeat(200_000)}]`;
		const t0 = performance.now();
		expect(bracketSpanEnd(big, 0)).toBe(big.length);
		expect(performance.now() - t0).toBeLessThan(500);
	});
});
