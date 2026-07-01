// Input model for the writer — the mirror of the reader's output types. A workbook is described as
// plain data (sheets of row-major value arrays); the writer infers each cell's OOXML type from the
// JS value, so the caller never spells out `t="..."` or number formats. This is the "value
// extractor, not object model" philosophy applied to writing.

/**
 * A value a cell can hold when writing. The OOXML cell type is inferred from it:
 * `string` → inline string, `number` → numeric, `boolean` → `b`, `Date` → date-styled serial.
 * `null` / `undefined` (including array holes) → an empty cell, omitted from the output.
 */
export type CellValue = string | number | boolean | Date | null | undefined

export interface SheetInput {
	/** Tab name. 1–31 chars, unique (case-insensitively), free of `\ / ? * [ ] :`. */
	readonly name: string
	/** Rows top-to-bottom; each a left-to-right array of cell values. `rows[0][0]` is A1. */
	readonly rows: readonly (readonly CellValue[])[]
}

export interface WorkbookInput {
	/** At least one sheet, in tab order. */
	readonly sheets: readonly SheetInput[]
}

export interface WriteOptions {
	/** Use the 1904 date epoch (legacy Mac) instead of the default 1900 system. */
	readonly date1904?: boolean
}
