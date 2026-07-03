import { describe, expect, it } from "vitest"
import { MAX_FORMULA_LEN, translateFormula } from "../formula"

// F5.4 — the shared-formula translator. The vectors below are openpyxl's own Translator output
// (openpyxl.formula.translate.Translator) for a master formula shifted by (Δrow, Δcol); matching
// them pins our translator to the reference behavior, including the subtle cases the scope calls
// out: cross-sheet cell parts DO shift (the sheet name stays), $-absolute parts pin, strings and
// function names are never touched.

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
]

describe("translateFormula — matches openpyxl's Translator", () => {
	it("reproduces every reference vector", () => {
		for (const [formula, dRow, dCol, expected] of VECTORS) {
			expect(translateFormula(formula, dRow, dCol), formula).toBe(expected)
		}
	})

	it("is the identity for a zero offset (the master cell itself)", () => {
		expect(translateFormula("A1+$B$2*Sheet1!C3", 0, 0)).toBe("A1+$B$2*Sheet1!C3")
	})

	it("rewrites a reference shifted off the grid to #REF! (Excel behavior; openpyxl throws)", () => {
		expect(translateFormula("A1", -1, 0)).toBe("#REF!") // row 0
		expect(translateFormula("A1", 0, -1)).toBe("#REF!") // column 0
		expect(translateFormula("B2*2", -5, 0)).toBe("#REF!*2") // only the off-grid ref becomes #REF!
	})

	it("does not shift things that only look like references", () => {
		// A lowercase name (stored refs are uppercase), a function name, and a string are left alone.
		expect(translateFormula('myRange+SUM(A1)&"B7"', 1, 0)).toBe('myRange+SUM(A2)&"B7"')
		// A pseudo-column past XFD (ZZZ = 18278 > 16384) is not an addressable cell — leave verbatim.
		expect(translateFormula("ZZZ1+A1", 1, 0)).toBe("ZZZ1+A2")
	})

	it("exposes Excel's formula length ceiling", () => {
		expect(MAX_FORMULA_LEN).toBe(8192)
	})
})
