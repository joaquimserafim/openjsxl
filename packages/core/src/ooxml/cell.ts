import type { Cell } from '../types'

// Turn a worksheet `<c>` element's raw pieces into a typed Cell. The `t` attribute selects
// how the value text is interpreted; when absent the cell is a number (the OOXML default).
//
//   t          value source        meaning
//   ─────────  ──────────────────  ──────────────────────────────────────────────
//   (absent)   <v>                 number
//   "n"        <v>                 number
//   "s"        <v>                 zero-based index into the shared string table
//   "str"      <v>                 a formula's cached string result
//   "inlineStr" <is><t>            text stored inline on the cell
//   "b"        <v>                 boolean — "1"/"0" (NOT a number)
//   "e"        <v>                 error literal, e.g. "#DIV/0!"
//
// Booleans and errors live inside `<v>` and must not be read as numbers. Malformed or
// missing values degrade to an `empty` cell rather than throwing — a reader stays resilient
// on imperfect files, matching the tokenizer and zip layers. Date typing is deferred to
// F2.1 (a date-styled serial reads as a plain number until then), so the style index is not
// needed here yet.

export interface RawCell {
	/** A1 reference, e.g. "B2". */
	ref: string
	/** The `t` attribute; `undefined` means the default (number). */
	type: string | undefined
	/** `<v>` text, or concatenated inline `<is>` text; `undefined` when the cell has none. */
	value: string | undefined
}

export interface DecodeContext {
	/** The workbook's shared string table (F1.5), indexed by `s`-type cells. */
	sharedStrings: string[]
}

export function decodeCell(raw: RawCell, ctx: DecodeContext): Cell {
	const { ref, value } = raw

	switch (raw.type) {
		case 's': {
			// Shared string: the value is an index. Out-of-range or non-integer indices
			// point at nothing, so the cell reads as empty rather than inventing text.
			const index = Number.parseInt(value ?? '', 10)
			const resolved = Number.isInteger(index) ? ctx.sharedStrings[index] : undefined
			return resolved === undefined
				? { ref, type: 'empty', value: null }
				: { ref, type: 'string', value: resolved }
		}
		case 'inlineStr':
		case 'str':
			return value === undefined
				? { ref, type: 'empty', value: null }
				: { ref, type: 'string', value }
		case 'b':
			return value === undefined
				? { ref, type: 'empty', value: null }
				: { ref, type: 'boolean', value: value === '1' }
		case 'e':
			return value === undefined
				? { ref, type: 'empty', value: null }
				: { ref, type: 'error', value }
		default: {
			// Absent or "n": a number. Reject empty/non-finite content as empty.
			if (value === undefined || value === '') return { ref, type: 'empty', value: null }
			const num = Number(value)
			return Number.isFinite(num)
				? { ref, type: 'number', value: num }
				: { ref, type: 'empty', value: null }
		}
	}
}
