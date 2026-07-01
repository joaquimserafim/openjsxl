import type { Cell } from '../types'
import { serialToDate } from './dates'
import type { StyleTable } from './styles'

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
// on imperfect files, matching the tokenizer and zip layers. A numeric cell becomes a `date`
// only when the context carries a style table and the cell's style applies a date/time
// number format (F2.1); nothing about a number's value reveals that on its own.

export interface RawCell {
	/** A1 reference, e.g. "B2". */
	readonly ref: string
	/** The `t` attribute; `undefined` means the default (number). */
	readonly type: string | undefined
	/** `<v>` text, or concatenated inline `<is>` text; `undefined` when the cell has none. */
	readonly value: string | undefined
	/** The `s` attribute (index into `cellXfs`); drives date detection. `undefined` ⇒ style 0. */
	readonly style: number | undefined
}

export interface DecodeContext {
	/** The workbook's shared string table (F1.5), indexed by `s`-type cells. */
	readonly sharedStrings: readonly string[]
	/** Style table (F2.1); when present, date-styled numbers decode as `date` cells. */
	readonly styles?: StyleTable
	/** Workbook 1904 date system flag, selecting the serial epoch. Defaults to false. */
	readonly date1904?: boolean
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
			if (!Number.isFinite(num)) return { ref, type: 'empty', value: null }
			// A date-styled serial is a date — unless the serial is so large (or small) that it
			// falls outside JS's representable Date range, where serialToDate yields an Invalid Date
			// (getTime() === NaN). A broken Date helps no consumer (and would crash the writer's
			// dateToSerial on a round trip), so such a cell stays a plain number.
			if (ctx.styles?.isDateStyle(raw.style)) {
				const date = serialToDate(num, ctx.date1904 ?? false)
				if (!Number.isNaN(date.getTime())) return { ref, type: 'date', value: date }
			}
			return { ref, type: 'number', value: num }
		}
	}
}
