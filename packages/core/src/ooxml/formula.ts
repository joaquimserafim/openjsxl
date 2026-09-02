import { columnToIndex, indexToColumn, MAX_COL, MAX_ROW } from "./a1";

// Formula text (F5.4) — no evaluation, just fidelity. Two jobs live here: the length ceiling the
// writer enforces, and the shared-formula translator the reader needs.
//
// A SHARED formula stores its text once on the master cell (`<f t="shared" ref si>text</f>`); every
// dependent cell carries only `<f t="shared" si/>` and is understood to be the master's text with
// its RELATIVE references shifted by the dependent's offset from the master. To hand a caller the
// real formula for a dependent, we translate: shift the relative parts of each same-sheet reference
// by (Δrow, Δcol), leave `$`-absolute parts pinned, and leave everything else — strings, function
// names, cross-sheet references, defined names, structured references — exactly as written. This is
// openpyxl's `Translator` behavior. Correctness rests on TOKENIZING rather than blind regex
// replacement: a naive search-and-replace would corrupt `"A1"` inside a string, `LOG10(` as a
// function, `Sheet1!A1`, or the `A1` column name inside `Table1[A1]`.
//
// F11.9: references are recognized CASE-INSENSITIVELY (Excel accepts `a1`; a producer may store it)
// and every reference we rewrite is emitted in canonical uppercase. A bracketed span — a structured
// reference (`Table1[Col]`, `Table1[[#Data],[A1]]`, `[@Col]`) or an external-workbook index
// (`[1]Sheet1!A1`) — is OPAQUE: it is copied verbatim, so A1-looking column names never shift.
// Inside brackets a `'` escapes the next character (`[Col'[1']]`), and nesting is tracked. An
// unbalanced `[` is ambiguous text: from that point on the formula is copied unchanged (never a
// partial rewrite, never a throw). A reference must start at a formula boundary (`2e5` is a number,
// `ñA1` a name) and its row is 1-based (`A0`, `0:0`, `A01` are names, as in Excel/openpyxl).
// Translation is a single left-to-right pass — linear in the text.

/** Excel's hard ceiling on formula text length (characters). The writer refuses anything longer. */
export const MAX_FORMULA_LEN = 8192;

// One token at a time, left to right (sticky). Order matters: a string or quoted sheet name is
// consumed whole before its contents can be mistaken for references; a WHOLE-column range (`A:A`) and
// WHOLE-row range (`1:1`) are recognized before a bare column/row could be mistaken for a name/number
// (they only exist adjacent to `:`); a single reference is tried before a bare identifier so `A1`
// reads as a cell, not a name; the final `[\s\S]` guarantees progress on any operator/paren/comma.
// A trailing boundary keeps `A1` in `A1B`, `A1_x`, or `A1.5` — or `A:A` in `A:A10` — part of the
// larger token. `[` is handled by the scanner before this regex runs, so it never reaches it as the
// start of a token.
const TOKEN =
	/("(?:[^"]|"")*")|('(?:[^']|'')*')|(\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![0-9A-Za-z_.])|(\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6})(?![0-9A-Za-z_.])|(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?![0-9A-Za-z_.])|([A-Za-z_][A-Za-z0-9_.]*)|(\s+)|([\s\S])/gy;

const REF = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/;
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/;
const ROW_RANGE = /^(\$?)([1-9][0-9]{0,6}):(\$?)([1-9][0-9]{0,6})$/;

// A reference can only START at a formula boundary. The sticky TOKEN regex carries a trailing
// boundary, but after a one-char fallback token (`2` of `2e5`, the `ñ` of a Unicode-prefixed name)
// the scanner restarts at the next position, where `e5` / `A1` would otherwise match as a cell.
// So a col-range/row-range/ref token is only a reference when the character before it is absent or
// is not identifier-ish: not a digit, ASCII letter, `_`, `.`, or any non-ASCII code unit. (`$` is
// part of the token itself, so `$A1` is judged by the character before the `$`.)
/** @internal Exported for direct unit tests only — not part of the public API. */
export function refCanStartAfter(code: number): boolean {
	if (code < 0) return true;
	if (code >= 0x80) return false;
	if (code >= 0x30 && code <= 0x39) return false; // 0-9
	if (code >= 0x41 && code <= 0x5a) return false; // A-Z
	if (code >= 0x61 && code <= 0x7a) return false; // a-z
	return code !== 0x5f && code !== 0x2e; // _ .
}

// Shift one already-matched reference by the delta, honoring `$` pins. Returns the rewritten ref in
// canonical uppercase, or `undefined` when the token isn't actually an addressable cell (a 3-letter
// pseudo-column past XFD, e.g. the `ZZZ` of a name), so the caller leaves it verbatim. A shift that
// lands off the grid becomes `#REF!`, exactly as Excel rewrites a reference it can no longer address.
/** @internal Exported for direct unit tests only — not part of the public API. */
export function shiftRef(ref: string, dRow: number, dCol: number): string | undefined {
	const m = REF.exec(ref);
	if (m === null) return undefined;
	const colAbs = m[1] as string;
	const rowAbs = m[3] as string;
	let col = columnToIndex(m[2] as string);
	let row = Number.parseInt(m[4] as string, 10);
	if (col > MAX_COL || row > MAX_ROW) return undefined; // not a real cell → not a reference
	if (colAbs !== "$") col += dCol;
	if (rowAbs !== "$") row += dRow;
	if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return "#REF!";
	return `${colAbs}${indexToColumn(col)}${rowAbs}${row}`;
}

// A whole-column range (`A:C`, `$A:$B`, or the cell part of `Sheet1!A:C`): shift each relative
// column endpoint. Endpoints past XFD aren't real columns, so the token is left verbatim; a shift off
// the grid becomes `#REF!`. Whole-row ranges (`1:1`, `$2:$5`) are the same with rows.
/** @internal Exported for direct unit tests only — not part of the public API. */
export function shiftColRange(token: string, dCol: number): string {
	const m = COL_RANGE.exec(token);
	if (m === null) return token;
	const i1 = columnToIndex(m[2] as string);
	const i2 = columnToIndex(m[4] as string);
	if (i1 > MAX_COL || i2 > MAX_COL) return token;
	const n1 = m[1] === "$" ? i1 : i1 + dCol;
	const n2 = m[3] === "$" ? i2 : i2 + dCol;
	if (n1 < 1 || n1 > MAX_COL || n2 < 1 || n2 > MAX_COL) return "#REF!";
	return `${m[1]}${indexToColumn(n1)}:${m[3]}${indexToColumn(n2)}`;
}

/** @internal Exported for direct unit tests only — not part of the public API. */
export function shiftRowRange(token: string, dRow: number): string {
	const m = ROW_RANGE.exec(token);
	if (m === null) return token;
	const r1 = Number.parseInt(m[2] as string, 10);
	const r2 = Number.parseInt(m[4] as string, 10);
	if (r1 > MAX_ROW || r2 > MAX_ROW) return token;
	const n1 = m[1] === "$" ? r1 : r1 + dRow;
	const n2 = m[3] === "$" ? r2 : r2 + dRow;
	if (n1 < 1 || n1 > MAX_ROW || n2 < 1 || n2 > MAX_ROW) return "#REF!";
	return `${m[1]}${n1}:${m[3]}${n2}`;
}

// The end (exclusive) of the balanced bracket span starting at `open` (which is `[`), or -1 when it
// never closes. Nesting counts (`[[#Data],[Col]]`); a `'` inside escapes the following character
// (`[Col'[1']]`, `[Col']']`). One pass over the span — linear.
/** @internal Exported for direct unit tests only — not part of the public API. */
export function bracketSpanEnd(formula: string, open: number): number {
	let depth = 0;
	for (let i = open; i < formula.length; i++) {
		const c = formula[i];
		if (c === "'") {
			i++; // the escaped character, whatever it is
		} else if (c === "[") {
			depth++;
		} else if (c === "]") {
			depth--;
			if (depth === 0) return i + 1;
		}
	}
	return -1;
}

/**
 * Translate a shared-formula master's text for a dependent cell offset (Δrow, Δcol) from the master.
 * Every relative cell reference shifts — including the cell part of a cross-sheet reference like
 * `Sheet1!A1` (the sheet name is a separate token and stays) — while `$`-absolute parts, string
 * literals, function/defined names, and bracketed spans (structured references, external-workbook
 * indexes) are left untouched, matching openpyxl's `Translator`. References are matched
 * case-insensitively and rewritten in canonical uppercase. A shift off the grid becomes `#REF!`.
 * An unbalanced `[` leaves the remainder of the text unchanged. A zero offset (the master cell
 * itself) returns the text unchanged.
 */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
	if (dRow === 0 && dCol === 0) return formula;
	let out = "";
	let pos = 0;
	while (pos < formula.length) {
		if (formula[pos] === "[") {
			const end = bracketSpanEnd(formula, pos);
			if (end === -1) return out + formula.slice(pos); // ambiguous: copy the rest unchanged
			out += formula.slice(pos, end);
			pos = end;
			continue;
		}
		TOKEN.lastIndex = pos;
		const m = TOKEN.exec(formula);
		if (m === null) return out + formula.slice(pos); // unreachable ([\s\S] always matches)
		const start = pos;
		pos = TOKEN.lastIndex;
		const colRange = m[3];
		const rowRange = m[4];
		const ref = m[5];
		const next = formula[pos];
		if (
			(colRange !== undefined || rowRange !== undefined || ref !== undefined) &&
			!refCanStartAfter(start > 0 ? formula.charCodeAt(start - 1) : -1)
		) {
			// `2e5`, `1E3`, `ñA1`: a numeral's exponent or the tail of a longer name, not a cell.
			out += m[0];
		} else if (colRange !== undefined) {
			// `A:A[` cannot start a structured reference, but keep the same guard shape as cells.
			out += next === "!" ? colRange : shiftColRange(colRange, dCol);
		} else if (rowRange !== undefined) {
			out += shiftRowRange(rowRange, dRow);
		} else if (ref !== undefined) {
			// A ref-shaped token is NOT a cell when it's a function name (`LOG10(`), an unquoted sheet
			// name (`ABC1!…`), or a table name (`A1[Col]`); those stay verbatim. Everything else shifts.
			out +=
				next === "!" || next === "(" || next === "["
					? ref
					: (shiftRef(ref, dRow, dCol) ?? ref);
		} else {
			// Strings, quoted sheet names, identifiers, whitespace, and operators pass through verbatim.
			out += m[0];
		}
	}
	return out;
}
