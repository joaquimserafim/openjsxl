import { columnToIndex, indexToColumn, MAX_COL, MAX_ROW } from "./a1"

// Formula text (F5.4) — no evaluation, just fidelity. Two jobs live here: the length ceiling the
// writer enforces, and the shared-formula translator the reader needs.
//
// A SHARED formula stores its text once on the master cell (`<f t="shared" ref si>text</f>`); every
// dependent cell carries only `<f t="shared" si/>` and is understood to be the master's text with
// its RELATIVE references shifted by the dependent's offset from the master. To hand a caller the
// real formula for a dependent, we translate: shift the relative parts of each same-sheet reference
// by (Δrow, Δcol), leave `$`-absolute parts pinned, and leave everything else — strings, function
// names, cross-sheet references, defined names — exactly as written. This is openpyxl's `Translator`
// behavior. Correctness rests on TOKENIZING rather than blind regex replacement: a naive
// search-and-replace would corrupt `"A1"` inside a string, `LOG10(` as a function, or `Sheet1!A1`.

/** Excel's hard ceiling on formula text length (characters). The writer refuses anything longer. */
export const MAX_FORMULA_LEN = 8192

// One token at a time, left to right (sticky). Order matters: a string or quoted sheet name is
// consumed whole before its contents can be mistaken for references; a WHOLE-column range (`A:A`) and
// WHOLE-row range (`1:1`) are recognized before a bare column/row could be mistaken for a name/number
// (they only exist adjacent to `:`); a single reference is tried before a bare identifier so `A1`
// reads as a cell, not a name; the final `[\s\S]` guarantees progress on any operator/paren/comma.
// References are uppercase-only (stored formulas normalize columns to uppercase) with a trailing
// boundary so `A1` in `A1B` or `A1_x` — or `A:A` in `A:A10` — stays part of the larger token.
const TOKEN =
	/("(?:[^"]|"")*")|('(?:[^']|'')*')|(\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![0-9A-Za-z_.])|(\$?[0-9]{1,7}:\$?[0-9]{1,7})(?![0-9A-Za-z_.])|(\$?[A-Z]{1,3}\$?[0-9]{1,7})(?![0-9A-Za-z_.])|([A-Za-z_][A-Za-z0-9_.]*)|(\s+)|([\s\S])/gy

const REF = /^(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})$/
const COL_RANGE = /^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/
const ROW_RANGE = /^(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})$/

// Shift one already-matched reference by the delta, honoring `$` pins. Returns the rewritten ref, or
// `undefined` when the token isn't actually an addressable cell (a 3-letter pseudo-column past XFD,
// e.g. the `ZZZ` of a name), so the caller leaves it verbatim. A shift that lands off the grid
// becomes `#REF!`, exactly as Excel rewrites a reference it can no longer address.
function shiftRef(ref: string, dRow: number, dCol: number): string | undefined {
	const m = REF.exec(ref)
	if (m === null) return undefined
	const colAbs = m[1] as string
	const rowAbs = m[3] as string
	let col = columnToIndex(m[2] as string)
	let row = Number.parseInt(m[4] as string, 10)
	if (col > MAX_COL || row > MAX_ROW) return undefined // not a real cell → not a reference
	if (colAbs !== "$") col += dCol
	if (rowAbs !== "$") row += dRow
	if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return "#REF!"
	return `${colAbs}${indexToColumn(col)}${rowAbs}${row}`
}

// A whole-column range (`A:C`, `$A:$B`, or the cell part of `Sheet1!A:C`): shift each relative
// column endpoint. Endpoints past XFD aren't real columns, so the token is left verbatim; a shift off
// the grid becomes `#REF!`. Whole-row ranges (`1:1`, `$2:$5`) are the same with rows.
function shiftColRange(token: string, dCol: number): string {
	const m = COL_RANGE.exec(token)
	if (m === null) return token
	const i1 = columnToIndex(m[2] as string)
	const i2 = columnToIndex(m[4] as string)
	if (i1 > MAX_COL || i2 > MAX_COL) return token
	const n1 = m[1] === "$" ? i1 : i1 + dCol
	const n2 = m[3] === "$" ? i2 : i2 + dCol
	if (n1 < 1 || n1 > MAX_COL || n2 < 1 || n2 > MAX_COL) return "#REF!"
	return `${m[1]}${indexToColumn(n1)}:${m[3]}${indexToColumn(n2)}`
}

function shiftRowRange(token: string, dRow: number): string {
	const m = ROW_RANGE.exec(token)
	if (m === null) return token
	const r1 = Number.parseInt(m[2] as string, 10)
	const r2 = Number.parseInt(m[4] as string, 10)
	if (r1 > MAX_ROW || r2 > MAX_ROW) return token
	const n1 = m[1] === "$" ? r1 : r1 + dRow
	const n2 = m[3] === "$" ? r2 : r2 + dRow
	if (n1 < 1 || n1 > MAX_ROW || n2 < 1 || n2 > MAX_ROW) return "#REF!"
	return `${m[1]}${n1}:${m[3]}${n2}`
}

/**
 * Translate a shared-formula master's text for a dependent cell offset (Δrow, Δcol) from the master.
 * Every relative cell reference shifts — including the cell part of a cross-sheet reference like
 * `Sheet1!A1` (the sheet name is a separate token and stays) — while `$`-absolute parts, string
 * literals, and function/defined names are left untouched, matching openpyxl's `Translator`. A shift
 * off the grid becomes `#REF!`. A zero offset (the master cell itself) returns the text unchanged.
 */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
	if (dRow === 0 && dCol === 0) return formula
	let out = ""
	TOKEN.lastIndex = 0
	let m: RegExpExecArray | null = TOKEN.exec(formula)
	while (m !== null) {
		const colRange = m[3]
		const rowRange = m[4]
		const ref = m[5]
		if (colRange !== undefined) {
			out += shiftColRange(colRange, dCol)
		} else if (rowRange !== undefined) {
			out += shiftRowRange(rowRange, dRow)
		} else if (ref !== undefined) {
			// A ref-shaped token is NOT a cell when it's a function name (`LOG10(`) or an unquoted
			// sheet name (`ABC1!…`); those stay verbatim. Everything else shifts.
			const next = formula[TOKEN.lastIndex]
			out += next === "!" || next === "(" ? ref : (shiftRef(ref, dRow, dCol) ?? ref)
		} else {
			// Strings, quoted sheet names, identifiers, whitespace, and operators pass through verbatim.
			out += m[0]
		}
		m = TOKEN.exec(formula)
	}
	return out
}
