import { describe, expect, it } from "vitest";
import { MAX_FORMULA_LEN, translateFormula } from "../formula";

// F5.4 — the shared-formula translator. The vectors below are openpyxl's own Translator output
// (openpyxl.formula.translate.Translator) for a master formula shifted by (Δrow, Δcol); matching
// them pins our translator to the reference behavior, including the subtle cases the scope calls
// out: cross-sheet cell parts DO shift (the sheet name stays), $-absolute parts pin, strings and
// function names are never touched.
//
// F11.9 adds: references are recognized case-insensitively and rewritten in canonical uppercase;
// bracketed spans (structured references, external-workbook indexes) are opaque; malformed brackets
// degrade to "leave the rest unchanged". Where we deliberately diverge from openpyxl it is noted.

// [master formula, Δrow, Δcol, expected]
const VECTORS: [string, number, number, string][] = [
	["A1*2", 2, 1, "B3*2"],
	["$A$1*2", 2, 1, "$A$1*2"],
	["$A1+A$1", 3, 2, "$A4+C$1"],
	["SUM(A1:A10)", 1, 0, "SUM(A2:A11)"],
	["SUM(B2:D2)", 3, 0, "SUM(B5:D5)"],
	["A1+B1+C1", 1, 0, "A2+B2+C2"],
	["Sheet2!A1+1", 1, 1, "Sheet2!B2+1"],
	["'My Sheet'!A1*2", 1, 1, "'My Sheet'!B2*2"],
	['"A1 literal"&A1', 1, 1, '"A1 literal"&B2'],
	["LOG10(A1)+LN(B2)", 1, 1, "LOG10(B2)+LN(C3)"],
	["A1&Sheet1!B2", 2, 2, "C3&Sheet1!D4"],
	["B2*C3", -1, -1, "A1*B2"],
	["IF(A1>0,B1,C1)", 4, 0, "IF(A5>0,B5,C5)"],
	["$B$2:$C$3", 2, 2, "$B$2:$C$3"],
	["AA10+AB11", 2, 2, "AC12+AD13"],
	["SUM(A1:A3)*Sheet2!B1", 0, 3, "SUM(D1:D3)*Sheet2!E1"],
	["A1<=B1", 99, 0, "A100<=B100"],
	["-A1", 1, 1, "-B2"],
	['CONCATENATE(A1," ",B1)', 8, 0, 'CONCATENATE(A9," ",B9)'],
	["$Z$100", 100, 1, "$Z$100"],
	['TEXT(A1,"0.00")', 0, 1, 'TEXT(B1,"0.00")'],
	// Whole-column and whole-row ranges (review regression — these were left unshifted).
	["SUM(A:A)", 0, 1, "SUM(B:B)"],
	["SUM(A:B)", 0, 1, "SUM(B:C)"],
	["SUM($A:$B)", 0, 2, "SUM($A:$B)"],
	["SUM(A:$B)", 0, 2, "SUM(C:$B)"],
	["1:1", 1, 0, "2:2"],
	["2:5", 2, 0, "4:7"],
	["SUM($2:$5)", 3, 0, "SUM($2:$5)"],
	["VLOOKUP(A1,Data!A:C,3,0)", 0, 1, "VLOOKUP(B1,Data!B:D,3,0)"],
	["A5:A", 0, 1, "B5:A"], // mixed: only the full-ref side shifts
	["A:A10", 0, 1, "A:B10"],
	["COUNTIF(B:B,A1)", 4, 2, "COUNTIF(D:D,C5)"],
];

// F11.9 — openpyxl Translator output (origin C1 → C2 is Δ(1,0); C1 → D3 is Δ(2,1)) for structured
// references, lowercase references, external-workbook brackets, and their look-alikes.
const F119_VECTORS: [string, number, number, string][] = [
	// Structured references are opaque — an A1-looking column name never shifts.
	["SUM(Table1[A1])", 1, 0, "SUM(Table1[A1])"],
	["SUM(Table1[A1])", 2, 1, "SUM(Table1[A1])"],
	["Table1[@A1]", 2, 1, "Table1[@A1]"],
	["Table1[[#Data],[A1]]", 2, 1, "Table1[[#Data],[A1]]"],
	["SUM(Table1[[#Headers],[Col'[1']]])", 2, 1, "SUM(Table1[[#Headers],[Col'[1']]])"],
	["[@A1]*2", 2, 1, "[@A1]*2"], // bare structured ref inside a table body
	["A1[Col]", 2, 1, "A1[Col]"], // an A1-looking TABLE name
	// Lowercase references shift and come back canonical uppercase (the F11.9 defect).
	["a1+Table1[B2]", 1, 0, "A2+Table1[B2]"],
	["a1+Table1[B2]", 2, 1, "B3+Table1[B2]"],
	["a1", 2, 1, "B3"],
	["a:a", 2, 1, "B:B"],
	["sum(a1:b2)", 2, 1, "sum(B3:C4)"], // the lowercase function NAME stays as written
	["log10(a1)", 2, 1, "log10(B3)"],
	["abc1!a1", 2, 1, "abc1!B3"], // unquoted sheet name stays; its cell part shifts
	["'my sheet'!a1", 2, 1, "'my sheet'!B3"],
	['"a1"&a1', 2, 1, '"a1"&B3'],
	// External-workbook brackets are opaque; the cell part still shifts (openpyxl agrees).
	["[1]Sheet1!A1", 2, 1, "[1]Sheet1!B3"],
	["[Book.xlsx]Sheet1!a1", 2, 1, "[Book.xlsx]Sheet1!B3"],
	// Boundaries: not references.
	["name_a1", 2, 1, "name_a1"],
	["A1_x", 2, 1, "A1_x"],
	["A1.5", 2, 1, "A1.5"],
	// A reference must START at a boundary (review A1): a numeral's exponent is not a cell, whether
	// lowercase (the F11.9 regression the review caught) or uppercase (pre-existing), and a
	// Unicode-prefixed name keeps its A1-shaped tail.
	["2e5+a1", 1, 0, "2e5+A2"],
	["2e5", 0, 1, "2e5"],
	["1E3+A1", 1, 0, "1E3+A2"],
	["1.5E+3*A1", 1, 1, "1.5E+3*B2"],
	["ñA1+A1", 1, 0, "ñA1+A2"],
	["日本A1", 1, 0, "日本A1"],
	// Rows are 1-based (review A2): `A0`, `0:0`, `A01` are names, never cells — verbatim, and a
	// fully pinned `$A$0` is byte-stable rather than `#REF!`.
	["A0+A1", 1, 0, "A0+A2"],
	["a0", 1, 0, "a0"],
	["$A$0", 1, 0, "$A$0"],
	["0:0", 1, 0, "0:0"],
	["A01+A1", 1, 0, "A01+A2"],
];

// Reopened F11.9: names and complete sheet qualifiers are opaque, including reference-shaped parts.
const IDENTIFIER_VECTORS: [string, number, number, string][] = [
	["a1é+A1", 1, 1, "a1é+B2"],
	["A1é+a1", 1, 1, "A1é+B2"],
	["éa1+a1", 1, 1, "éa1+B2"],
	["a1日本b2+A1", 1, 1, "a1日本b2+B2"],
	["a1\u0301+A1", 1, 1, "a1\u0301+B2"],
	["a1𐐀+A1", 1, 1, "a1𐐀+B2"],
	["𐐀a1+A1", 1, 1, "𐐀a1+B2"],
	["a1١+A1", 1, 1, "a1١+B2"],
	["\\A1+a1\\b2+a1?b2+A1", 1, 1, "\\A1+a1\\b2+a1?b2+B2"],
	["SUM(a1é[A1])+A1", 1, 1, "SUM(a1é[A1])+B2"],
	["a1é!a1", 1, 1, "a1é!B2"],
	["SUM(1:2!A1)", 1, 1, "SUM(1:2!B2)"],
	["SUM(a1:b2!C3)", 1, 1, "SUM(a1:b2!D4)"],
	["SUM(A1:B2!C3)", 1, 1, "SUM(A1:B2!D4)"],
	["SUM(a1:Sheet2!C3)", 1, 1, "SUM(a1:Sheet2!D4)"],
	["SUM(Sheet1:b2!C3)", 1, 1, "SUM(Sheet1:b2!D4)"],
	["S!A1:S!B2", 1, 1, "S!B2:S!C3"],
	["'S'!A1:S!B2", 1, 1, "'S'!B2:S!C3"],
	["S!A1:'S'!B2", 1, 1, "S!B2:'S'!C3"],
	["[1]S!A1:S!B2", 1, 1, "[1]S!B2:S!C3"],
	["SUM(S!A1:S!B2)", 1, 1, "SUM(S!B2:S!C3)"],
	["!A1:S!B2", 1, 1, "!B2:S!C3"],
	// Error endpoints are not sheet qualifiers. openpyxl rejects the forward #REF! range.
	["SUM(A1:#REF!)", 1, 1, "SUM(B2:#REF!)"],
	["SUM(A1:#ref!)+A1:#ReF!", 1, 1, "SUM(B2:#ref!)+B2:#ReF!"],
	["SUM(#REF!:A1)", 1, 1, "SUM(#REF!:B2)"],
	["S!A1:#REF!", 1, 1, "S!B2:#REF!"],
	// Other errors are not valid reference endpoints; preserve their tolerant handling too.
	["A1:#VALUE!+A1:#NUM!+A1:#NULL!", 1, 1, "B2:#VALUE!+B2:#NUM!+B2:#NULL!"],
	["A1:#DIV/0!", 1, 1, "B2:#DIV/0!"],
	["A1#Sheet:B2#Sheet!A1", 1, 1, "A1#Sheet:B2#Sheet!B2"],
	["A1#REF!A1", 1, 1, "A1#REF!B2"],
	["a1#REF:b2#VALUE!A1", 1, 1, "a1#REF:b2#VALUE!B2"],
	["SUM(a1é:b2日!C3)", 1, 1, "SUM(a1é:b2日!D4)"],
	["A:B!a1", 1, 1, "A:B!B2"],
	["1:2!A:A", 1, 1, "1:2!B:B"],
	["1:2!1:3", 1, 1, "1:2!2:4"],
	["1:2!$A$1+a1", 1, 1, "1:2!$A$1+B2"],
	["[1]a1:b2!C3", 1, 1, "[1]a1:b2!D4"],
	["[Book.xlsx]1:2!A1", 1, 1, "[Book.xlsx]1:2!B2"],
	["'a1:b2'!C3", 1, 1, "'a1:b2'!D4"],
	["'[Book.xlsx]a1:b2'!C3", 1, 1, "'[Book.xlsx]a1:b2'!D4"],
	["'a1''é:b2'!C3", 1, 1, "'a1''é:b2'!D4"],
	["a1~:b2`!C3", 1, 1, "a1~:b2`!D4"],
	["SUM(A1:B2 Sheet1!C3)", 1, 1, "SUM(B2:C3 Sheet1!D4)"],
	["SUM(A1:B2,1:2!C3)", 1, 1, "SUM(B2:C3,1:2!D4)"],
	["A1;1:2!B2", 1, 1, "B2;1:2!C3"],
	["A1\n1:2!B2", 1, 1, "B2\n1:2!C3"],
	["IF(A1>0,1:2!B2,C3)", 1, 1, "IF(B2>0,1:2!C3,D4)"],
	['"1:2!A1"&A1', 1, 1, '"1:2!A1"&B2'],
	["T[[#Data],[1:2!A1]]+A1", 1, 1, "T[[#Data],[1:2!A1]]+B2"],
	["LOG10(A1)+2e5+1E3", 1, 1, "LOG10(B2)+2e5+1E3"],
	["a1é+A1+T[1:2!B1", 1, 1, "a1é+B2+T[1:2!B1"],
	["a1é+SUM(1:2!A1)", 0, 0, "a1é+SUM(1:2!A1)"],
	["a1é+SUM(1:2!A1)", -1, -1, "a1é+SUM(1:2!#REF!)"],
];

describe("translateFormula — matches openpyxl's Translator", () => {
	it("reproduces every reference vector", () => {
		for (const [formula, dRow, dCol, expected] of VECTORS) {
			expect(translateFormula(formula, dRow, dCol), formula).toBe(expected);
		}
	});

	it("reproduces the F11.9 structured / lowercase / external vectors", () => {
		for (const [formula, dRow, dCol, expected] of F119_VECTORS) {
			expect(translateFormula(formula, dRow, dCol), formula).toBe(expected);
		}
	});

	it.each(IDENTIFIER_VECTORS)(
		"preserves identifiers and sheet qualifiers in %s at offset (%i, %i)",
		(formula, dRow, dCol, expected) => {
			expect(translateFormula(formula, dRow, dCol)).toBe(expected);
		},
	);

	it("canonicalizes the WHOLE rewritten token to uppercase (deliberate divergence from openpyxl)", () => {
		// openpyxl leaves a fully pinned `$a$1` lowercase and only uppercases the parts it moves; we
		// emit Excel's canonical spelling for every reference we rewrite. Same cell either way.
		expect(translateFormula("$a$1+a$1+$a1", 1, 0)).toBe("$A$1+A$1+$A2");
	});

	it("is the identity for a zero offset (the master cell itself)", () => {
		expect(translateFormula("A1+$B$2*Sheet1!C3", 0, 0)).toBe("A1+$B$2*Sheet1!C3");
		expect(translateFormula("a1+Table1[", 0, 0)).toBe("a1+Table1["); // even malformed text
	});

	it("rewrites a reference shifted off the grid to #REF! (Excel behavior; openpyxl throws)", () => {
		expect(translateFormula("A1", -1, 0)).toBe("#REF!"); // row 0
		expect(translateFormula("A1", 0, -1)).toBe("#REF!"); // column 0
		expect(translateFormula("B2*2", -5, 0)).toBe("#REF!*2"); // only the off-grid ref becomes #REF!
		expect(translateFormula("XFD1", 0, 1)).toBe("#REF!"); // past the last column (openpyxl: XFE1)
	});

	it("does not shift things that only look like references", () => {
		// A real defined name, a function name, and a string are left alone; lowercase `a1` is a cell.
		expect(translateFormula('myRange+SUM(A1)&"B7"', 1, 0)).toBe('myRange+SUM(A2)&"B7"');
		// A pseudo-column past XFD (ZZZ = 18278 > 16384) is not an addressable cell — leave verbatim.
		expect(translateFormula("ZZZ1+A1", 1, 0)).toBe("ZZZ1+A2");
		expect(translateFormula("zzz1+a1", 1, 0)).toBe("zzz1+A2");
	});

	it("leaves everything from an unbalanced bracket onward unchanged (no partial rewrite, no throw)", () => {
		// openpyxl raises TokenizerError on an unmatched `[`; the tolerant reader must not.
		expect(translateFormula("Table1[", 1, 0)).toBe("Table1[");
		expect(translateFormula("A1+Table1[Col", 1, 0)).toBe("A2+Table1[Col"); // A1 before it still shifts
		expect(translateFormula("A1+Table1[[#Data],[Col]+B1", 1, 0)).toBe(
			"A2+Table1[[#Data],[Col]+B1",
		);
		// Divergence from openpyxl (its bracket matcher counts raw `[`/`]` and ignores `'`): we follow
		// Excel's structured-reference escape rule, where `'` escapes the next character inside a
		// header (`[Col']]` is the column named `Col]`), so `[Col']` is unclosed — the rest is verbatim.
		expect(translateFormula("A1+Table1[Col']+B1", 1, 0)).toBe("A2+Table1[Col']+B1");
		// Divergence from openpyxl (it leaves `]A1` whole): a stray `]` opens no span, so the `A1`
		// after it is an ordinary reference and shifts.
		expect(translateFormula("]A1", 1, 0)).toBe("]A2");
	});

	it("keeps balanced nested and escaped brackets byte-verbatim, and shifts what follows", () => {
		expect(translateFormula("Table1[[#Data],[Col']']]]+A1", 1, 0)).toBe(
			"Table1[[#Data],[Col']']]]+A2",
		);
		expect(translateFormula('T[[#This Row],[Item]]&"_x0041_"+a1', 1, 0)).toBe(
			'T[[#This Row],[Item]]&"_x0041_"+A2',
		);
	});

	it("stays linear on hostile input (a 1 MB structured reference; a 1 MB unbalanced bracket)", () => {
		const cell = "SUM(Table1[[#Data],[A1]])+";
		const big = cell.repeat(Math.ceil(1_000_000 / cell.length));
		const t0 = performance.now();
		const out = translateFormula(big, 1, 1);
		const balanced = performance.now() - t0;
		expect(out).toBe(big); // nothing to shift inside the spans; the `+` glue is verbatim
		const open = `A1+Table1[${"[x],".repeat(250_000)}`; // never closes
		const t1 = performance.now();
		expect(translateFormula(open, 1, 1)).toBe(`B2+Table1[${"[x],".repeat(250_000)}`);
		const unbalanced = performance.now() - t1;
		expect(balanced).toBeLessThan(2_000);
		expect(unbalanced).toBeLessThan(2_000);
	});

	it("exposes Excel's formula length ceiling", () => {
		expect(MAX_FORMULA_LEN).toBe(8192);
	});

	it("bounds failed qualifier lookahead across long runs without a bang", () => {
		// Each failed scan must be reused while the ordinary tokenizer walks the same run.
		// Digits and dollar signs take its single-code-unit fallback, exposing quadratic rescans.
		for (const prefix of ["9".repeat(500_000), "$".repeat(500_000), "a1é".repeat(170_000)]) {
			const formula = `${prefix}+A1`;
			const t0 = performance.now();
			expect(translateFormula(formula, 1, 1)).toBe(`${prefix}+B2`);
			expect(performance.now() - t0).toBeLessThan(2_000);
		}
		const formula = "SUM(1:2!a1)+".repeat(30_000);
		const t0 = performance.now();
		expect(translateFormula(formula, 1, 1)).toBe("SUM(1:2!B2)+".repeat(30_000));
		expect(performance.now() - t0).toBeLessThan(2_000);
	});
});
