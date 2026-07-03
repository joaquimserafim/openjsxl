// Input model for the writer — the mirror of the reader's output types. A workbook is described as
// plain data (sheets of row-major value arrays); the writer infers each cell's OOXML type from the
// JS value, so the caller never spells out `t="..."` or number formats. This is the "value
// extractor, not object model" philosophy applied to writing.

import type {
	CellStyle,
	ColumnProps,
	Comment,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetState,
} from "../types"

/**
 * A value a cell can hold when writing. The OOXML cell type is inferred from it:
 * `string` → inline string, `number` → numeric, `boolean` → `b`, `Date` → date-styled serial.
 * `null` / `undefined` (including array holes) → an empty cell, omitted from the output.
 */
export type CellValue = string | number | boolean | Date | null | undefined

/**
 * A cell with a {@link CellStyle} attached (F4.2). `value` is required but nullable — a styled
 * BLANK cell (`{ value: null, style }`) is real and emits `<c r s/>`, which is how a border or
 * fill lands on a cell with nothing in it. The style type is exactly what `Worksheet.style(ref)`
 * returns, so read → modify → write carries styles as a pass-through.
 */
export interface StyledCell {
	readonly value: CellValue
	readonly style?: CellStyle
}

/**
 * One cell in a row: a bare value, or a value with a style. Discrimination is total — `null` /
 * `undefined` mean empty, `Date` instances are dates, and any OTHER object must be a
 * {@link StyledCell} (an object without a `value` property throws `invalid-input`, catching stray
 * objects loudly). Bare-value rows are untouched: pre-F4.2 input keeps its exact meaning AND its
 * exact output bytes.
 */
export type CellInput = CellValue | StyledCell

export interface SheetInput {
	/** Tab name. 1–31 chars, unique (case-insensitively), free of `\ / ? * [ ] :`. */
	readonly name: string
	/**
	 * Rows top-to-bottom; each a left-to-right array of cells. `rows[0][0]` is A1. A hole or
	 * `undefined` row is an empty row (the bridge produces sparse arrays; hand-written input may
	 * too) — matching how `null`/`undefined`/holes inside a row mean empty cells.
	 */
	readonly rows: readonly (readonly CellInput[] | undefined)[]
	/**
	 * Column width/visibility declarations (F4.5) — the same shape `Worksheet.columns` returns.
	 * Ranges are 1-based and inclusive; entries need a `width` and/or `hidden: true`.
	 */
	readonly columns?: readonly ColumnProps[]
	/**
	 * Per-row height/visibility keyed by 1-based row index (F4.5) — the same records
	 * `Worksheet.rowProperties` holds. A row may have properties without having any cells.
	 */
	readonly rowProperties?: Readonly<Record<number, RowProps>>
	/** Freeze the top `rows` rows and/or leftmost `cols` columns (F4.5). */
	readonly freeze?: FreezePane
	/**
	 * Merged-cell ranges in canonical A1 form, top-left:bottom-right (e.g. `"A1:B2"`) — the same
	 * strings `Worksheet.mergedCells` returns (F4.6). Malformed, single-cell, out-of-grid, and
	 * overlapping ranges are rejected: Excel repair-prompts on them.
	 */
	readonly merges?: readonly string[]
	/**
	 * Hyperlinks on this sheet (F4.6) — the same records `Worksheet.hyperlinks` returns. Each needs
	 * a `ref` (cell or range) plus an external `target` and/or an in-workbook `location`; `tooltip`
	 * and `display` are optional. External targets get the sheet's relationships part.
	 */
	readonly hyperlinks?: readonly Hyperlink[]
	/**
	 * Tab visibility (F4.6). Defaults to `"visible"`; at least one sheet in the workbook must
	 * remain visible, or Excel refuses the file.
	 */
	readonly state?: SheetState
	/**
	 * Cell comments on this sheet (F5.2) — the same records `Worksheet.comments` returns. Each needs
	 * a single-cell `ref` and `text`; `author` is optional. Excel shows a comment only alongside a
	 * legacy VML drawing, which the writer emits automatically for every commented sheet.
	 */
	readonly comments?: readonly Comment[]
}

export interface WorkbookInput {
	/** At least one sheet, in tab order. */
	readonly sheets: readonly SheetInput[]
	/**
	 * The workbook's raw theme part (`xl/theme/theme1.xml`), carried verbatim (F5.3). When a written
	 * style uses a theme color, the writer emits this instead of the built-in Office theme — so a
	 * custom-theme file keeps its exact colors on rewrite. Must be a non-empty, XML-safe string;
	 * ignored when no written style needs a theme part.
	 */
	readonly themeXml?: string
}

export interface WriteOptions {
	/** Use the 1904 date epoch (legacy Mac) instead of the default 1900 system. */
	readonly date1904?: boolean
}
