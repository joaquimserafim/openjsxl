// Input model for the writer — the mirror of the reader's output types. A workbook is described as
// plain data (sheets of row-major value arrays); the writer infers each cell's OOXML type from the
// JS value, so the caller never spells out `t="..."` or number formats. This is the "value
// extractor, not object model" philosophy applied to writing.

import type {
	CellStyle,
	ColumnProps,
	Comment,
	DataValidation,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetImage,
	SheetState,
	TableInfo,
} from "../types";

/**
 * A value a cell can hold when writing. The OOXML cell type is inferred from it:
 * `string` → inline string, `number` → numeric, `boolean` → `b`, `Date` → date-styled serial.
 * `null` / `undefined` (including array holes) → an empty cell, omitted from the output.
 */
export type CellValue = string | number | boolean | Date | null | undefined;

/**
 * A cell carrying a {@link CellStyle} and/or a formula (F4.2, F5.4). Without a formula, `value` is
 * required but nullable — a styled BLANK cell (`{ value: null, style }`) is real and emits
 * `<c r s/>`, which is how a border or fill lands on an empty cell. With a `formula` (stored form,
 * no leading `=`), `value` is the optional cached result and the cell emits `<c s?><f>…</f><v>…</v></c>`.
 * The style type is exactly what `Worksheet.style(ref)` returns and `formula` is what
 * `Worksheet.formula(ref)` returns, so read → modify → write carries both as a pass-through.
 */
export interface StyledCell {
	readonly value?: CellValue;
	readonly style?: CellStyle;
	readonly formula?: string;
}

/**
 * One cell in a row: a bare value, or an object with a style and/or a formula. Discrimination is
 * total — `null` / `undefined` mean empty, `Date` instances are dates, and any OTHER object must be
 * a {@link StyledCell} (an object with neither a `value` property nor a `formula` throws
 * `invalid-input`, catching stray objects loudly). Bare-value rows are untouched: pre-F4.2 input
 * keeps its exact meaning AND its exact output bytes.
 */
export type CellInput = CellValue | StyledCell;

export interface SheetInput {
	/** Tab name. 1–31 chars, unique (case-insensitively), free of `\ / ? * [ ] :`. */
	readonly name: string;
	/**
	 * Rows top-to-bottom; each a left-to-right array of cells. `rows[0][0]` is A1. A hole or
	 * `undefined` row is an empty row (the bridge produces sparse arrays; hand-written input may
	 * too) — matching how `null`/`undefined`/holes inside a row mean empty cells.
	 */
	readonly rows: readonly (readonly CellInput[] | undefined)[];
	/**
	 * Column width/visibility declarations (F4.5) — the same shape `Worksheet.columns` returns.
	 * Ranges are 1-based and inclusive; entries need a `width` and/or `hidden: true`.
	 */
	readonly columns?: readonly ColumnProps[];
	/**
	 * Per-row height/visibility keyed by 1-based row index (F4.5) — the same records
	 * `Worksheet.rowProperties` holds. A row may have properties without having any cells.
	 */
	readonly rowProperties?: Readonly<Record<number, RowProps>>;
	/** Freeze the top `rows` rows and/or leftmost `cols` columns (F4.5). */
	readonly freeze?: FreezePane;
	/**
	 * Merged-cell ranges in canonical A1 form, top-left:bottom-right (e.g. `"A1:B2"`) — the same
	 * strings `Worksheet.mergedCells` returns (F4.6). Malformed, single-cell, out-of-grid, and
	 * overlapping ranges are rejected: Excel repair-prompts on them.
	 */
	readonly merges?: readonly string[];
	/**
	 * Hyperlinks on this sheet (F4.6) — the same records `Worksheet.hyperlinks` returns. Each needs
	 * a `ref` (cell or range) plus an external `target` and/or an in-workbook `location`; `tooltip`
	 * and `display` are optional. External targets get the sheet's relationships part.
	 */
	readonly hyperlinks?: readonly Hyperlink[];
	/**
	 * Tab visibility (F4.6). Defaults to `"visible"`; at least one sheet in the workbook must
	 * remain visible, or Excel refuses the file.
	 */
	readonly state?: SheetState;
	/**
	 * Cell comments on this sheet (F5.2) — the same records `Worksheet.comments` returns. Each needs
	 * a single-cell `ref` and `text`; `author` is optional. Excel shows a comment only alongside a
	 * legacy VML drawing, which the writer emits automatically for every commented sheet.
	 */
	readonly comments?: readonly Comment[];
	/**
	 * Pictures anchored on this sheet (F6.3) — the same records `Worksheet.images()` returns. Each
	 * needs an `anchor` (a one-cell `{from, ext}` or two-cell `{from, to}` cell anchor), the raw
	 * image `bytes`, and a `mime` (`image/png` | `image/jpeg` | `image/gif`); `name` is optional.
	 * Identical bytes are written once as a shared media part; EMU offsets/extents are used verbatim.
	 */
	readonly images?: readonly SheetImage[];
	/**
	 * Defined tables on this sheet (F9.1) — the same records `Worksheet.tables` returns. Column names
	 * DERIVE from the header row (the single source of truth), the numeric id is auto-assigned, and
	 * `name` must be a workbook-unique identifier (no spaces, not a cell reference). A table whose
	 * `ref` overlaps another, or whose header cells aren't non-empty text, is rejected.
	 */
	readonly tables?: readonly TableInfo[];
	/**
	 * Data-validation rules on this sheet (F9.2) — the same records `Worksheet.dataValidations`
	 * returns. Each needs a non-empty `sqref` (A1 ranges) and a `type`; `formula1`/`formula2` are
	 * operand text carried verbatim (a leading `=` is stripped). Prompt/error titles are capped at 32
	 * characters and bodies at 255; `showDropDown` uses the intuitive sense (`true` = arrow shown).
	 */
	readonly dataValidations?: readonly DataValidation[];
}

export interface WorkbookInput {
	/** At least one sheet, in tab order. */
	readonly sheets: readonly SheetInput[];
	/**
	 * The workbook's raw theme part (`xl/theme/theme1.xml`), carried verbatim (F5.3). When a written
	 * style uses a theme color, the writer emits this instead of the built-in Office theme — so a
	 * custom-theme file keeps its exact colors on rewrite. Must be a non-empty, XML-safe string;
	 * ignored when no written style needs a theme part.
	 */
	readonly themeXml?: string;
}

/**
 * Rows for the streaming writer (F5.1): a synchronous or asynchronous iterable of row arrays. Each
 * pulled item is one row in order (row 1, 2, 3, …) — yield an empty array or `undefined` for a blank
 * row. An `AsyncIterable` (a DB cursor, a paged fetch) is pulled only as the output is consumed, so a
 * slow source is never outpaced.
 */
export type StreamRows =
	| Iterable<readonly CellInput[] | undefined>
	| AsyncIterable<readonly CellInput[] | undefined>;

/**
 * One sheet for {@link streamXlsx} — the same shape as {@link SheetInput} except `rows` is a
 * {@link StreamRows} iterable rather than a materialized array, so a large sheet never lives in
 * memory all at once. Geometry and metadata (columns, rowProperties, freeze, merges, hyperlinks,
 * state, comments) stay upfront values, validated and emitted around the row stream. One caveat:
 * a row's height/hidden property applies to the row at that 1-based stream position, and a
 * rowProperties entry past the last streamed row is dropped (there is no row to attach it to).
 */
export interface StreamSheetInput {
	readonly name: string;
	readonly rows: StreamRows;
	readonly columns?: readonly ColumnProps[];
	readonly rowProperties?: Readonly<Record<number, RowProps>>;
	readonly freeze?: FreezePane;
	readonly merges?: readonly string[];
	readonly hyperlinks?: readonly Hyperlink[];
	readonly state?: SheetState;
	readonly comments?: readonly Comment[];
	/** Pictures anchored on this sheet (F6.3) — see {@link SheetInput.images}. */
	readonly images?: readonly SheetImage[];
	/**
	 * Defined tables on this sheet (F9.1) — see {@link SheetInput.tables}. The streaming writer can't
	 * read the header row upfront, so each table's column names come from `columns[i].name` here rather
	 * than being derived from the header cells.
	 */
	readonly tables?: readonly TableInfo[];
	/** Data-validation rules on this sheet (F9.2) — see {@link SheetInput.dataValidations}. */
	readonly dataValidations?: readonly DataValidation[];
}

/** A workbook for {@link streamXlsx}: sheets with streaming rows, plus the optional carried theme. */
export interface StreamWorkbookInput {
	readonly sheets: readonly StreamSheetInput[];
	readonly themeXml?: string;
}

export interface WriteOptions {
	/** Use the 1904 date epoch (legacy Mac) instead of the default 1900 system. */
	readonly date1904?: boolean;
}
