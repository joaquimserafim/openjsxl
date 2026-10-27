import type { Cell, DxfStyle } from "../types";
import { parseIsoDate, serialToDate } from "./dates";
import type { StyleTable } from "./styles";

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
//   "d"        <v>                 ISO 8601 date/time text, e.g. "2024-03-15T13:45:30" (F11.60)
//
// Booleans and errors live inside `<v>` and must not be read as numbers. Malformed or
// missing values degrade to an `empty` cell rather than throwing — a reader stays resilient
// on imperfect files, matching the tokenizer and zip layers. A numeric cell becomes a `date`
// only when the context carries a style table and the cell's style applies a date/time
// number format (F2.1); nothing about a number's value reveals that on its own.

export interface RawCell {
	/** A1 reference, e.g. "B2". */
	readonly ref: string;
	/** The `t` attribute; `undefined` means the default (number). */
	readonly type: string | undefined;
	/** `<v>` text, or concatenated inline `<is>` text; `undefined` when the cell has none. */
	readonly value: string | undefined;
	/** The `s` attribute (index into `cellXfs`); drives date detection. `undefined` ⇒ style 0. */
	readonly style: number | undefined;
}

export interface DecodeContext {
	/** The workbook's shared string table (F1.5), indexed by `s`-type cells. */
	readonly sharedStrings: readonly string[];
	/** Style table (F2.1); when present, date-styled numbers decode as `date` cells. */
	readonly styles?: StyleTable;
	/** Workbook 1904 date system flag, selecting the serial epoch. Defaults to false. */
	readonly date1904?: boolean;
	/**
	 * The workbook's `<dxfs>` table (F9.3), positional by `dxfId`. A conditional-formatting rule
	 * resolves its `@dxfId` against this to an inline {@link DxfStyle}. Absent when the workbook has
	 * no differential styles.
	 */
	readonly dxfs?: readonly DxfStyle[];
}

/**
 * Decode one `<c>` element's raw pieces into a typed Cell, by its `t` attribute (see the table above).
 * A value that is missing or does not parse reads as an empty cell.
 *
 * In each example `raw` is `{ ref: "A1", style: undefined }` plus the fields shown, and `ctx` is
 * `{ sharedStrings: ["x"] }` unless a style table is named.
 *
 * @example decodeCell({ ...raw, type: "s", value: "0" }, ctx) // { ref: "A1", type: "string", value: "x" }
 * @example decodeCell({ ...raw, type: "s", value: "9" }, ctx) // { ref: "A1", type: "empty", value: null } (no such string)
 * @example decodeCell({ ...raw, type: "str", value: "y" }, ctx) // { ref: "A1", type: "string", value: "y" }
 * @example decodeCell({ ...raw, type: "inlineStr", value: undefined }, ctx) // { ref: "A1", type: "empty", value: null }
 * @example decodeCell({ ...raw, type: "b", value: "1" }, ctx) // { ref: "A1", type: "boolean", value: true }
 * @example decodeCell({ ...raw, type: "b", value: undefined }, ctx) // { ref: "A1", type: "empty", value: null }
 * @example decodeCell({ ...raw, type: "e", value: "#N/A" }, ctx) // { ref: "A1", type: "error", value: "#N/A" }
 * @example decodeCell({ ...raw, type: "e", value: undefined }, ctx) // { ref: "A1", type: "empty", value: null }
 * @example decodeCell({ ...raw, type: "d", value: "2024-03-15" }, ctx) // { ref: "A1", type: "date", value: 2024-03-15T00:00:00.000Z }
 * @example decodeCell({ ...raw, type: "d", value: "2024-02-30" }, ctx) // { ref: "A1", type: "empty", value: null }
 * @example decodeCell({ ...raw, type: "n", value: "" }, ctx) // { ref: "A1", type: "empty", value: null }
 * @example decodeCell({ ...raw, type: "n", value: "abc" }, ctx) // { ref: "A1", type: "empty", value: null } (not a finite number)
 * @example decodeCell({ ...raw, value: "45366", style: 1 }, { ...ctx, styles: xf1IsADate }) // { ref: "A1", type: "date", value: 2024-03-15T00:00:00.000Z }
 * @example decodeCell({ ...raw, value: "1e300", style: 1 }, { ...ctx, styles: xf1IsADate }) // { ref: "A1", type: "number", value: 1e300 } (past Date's range)
 * @example decodeCell({ ...raw, type: "n", value: "42" }, ctx) // { ref: "A1", type: "number", value: 42 }
 */
export function decodeCell(raw: RawCell, ctx: DecodeContext): Cell {
	const { ref, value } = raw;

	switch (raw.type) {
		case "s": {
			// Shared string: the value is an index. Out-of-range or non-integer indices
			// point at nothing, so the cell reads as empty rather than inventing text.
			const index = Number.parseInt(value ?? "", 10);
			const resolved = Number.isInteger(index) ? ctx.sharedStrings[index] : undefined;
			return resolved === undefined
				? { ref, type: "empty", value: null }
				: { ref, type: "string", value: resolved };
		}
		case "inlineStr":
		case "str":
			return value === undefined
				? { ref, type: "empty", value: null }
				: { ref, type: "string", value };
		case "b":
			return value === undefined
				? { ref, type: "empty", value: null }
				: { ref, type: "boolean", value: value === "1" };
		case "e":
			return value === undefined
				? { ref, type: "empty", value: null }
				: { ref, type: "error", value };
		case "d": {
			// An ISO 8601 date, whatever the cell's style or the date system (F11.60). Text that is not
			// a real date reads as empty, never as a guessed date.
			const date = value === undefined ? undefined : parseIsoDate(value);
			return date === undefined
				? { ref, type: "empty", value: null }
				: { ref, type: "date", value: date };
		}
		default: {
			// Absent or "n": a number. Reject empty/non-finite content as empty.
			if (value === undefined || value === "") return { ref, type: "empty", value: null };
			const num = Number(value);
			if (!Number.isFinite(num)) return { ref, type: "empty", value: null };
			// A date-styled serial is a date — unless the serial is so large (or small) that it
			// falls outside JS's representable Date range, where serialToDate yields an Invalid Date
			// (getTime() === NaN). A broken Date helps no consumer (and would crash the writer's
			// dateToSerial on a round trip), so such a cell stays a plain number.
			if (ctx.styles?.isDateStyle(raw.style)) {
				const date = serialToDate(num, ctx.date1904 ?? false);
				if (!Number.isNaN(date.getTime())) return { ref, type: "date", value: date };
			}
			return { ref, type: "number", value: num };
		}
	}
}
