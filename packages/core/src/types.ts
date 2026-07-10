// Public cell model. A discriminated union, so narrowing on `type` also narrows
// `value`: `if (cell.type === 'date') { cell.value /* : Date */ }`.

export type CellType = "empty" | "string" | "number" | "boolean" | "date" | "error";

interface CellBase {
	/** A1 reference, e.g. "B2". */
	readonly ref: string;
}

export type Cell =
	| (CellBase & { readonly type: "empty"; readonly value: null })
	| (CellBase & { readonly type: "string"; readonly value: string })
	| (CellBase & { readonly type: "number"; readonly value: number })
	| (CellBase & { readonly type: "boolean"; readonly value: boolean })
	| (CellBase & { readonly type: "date"; readonly value: Date })
	| (CellBase & { readonly type: "error"; readonly value: string });

/** A populated row from {@link Worksheet.rows}. Sparse: absent rows/cells are simply omitted. */
export interface Row {
	/** 1-based row index — from `<row r>`, or positional when the attribute is absent. */
	readonly index: number;
	/** Cells present in the row, in document order. Gaps are simply absent (sparse). */
	readonly cells: readonly Cell[];
}

/**
 * A sheet tab's visibility (the `state` attribute on `<sheet>`). `hidden` sheets can be re-shown
 * from Excel's UI; `veryHidden` ones only through VBA or by editing the file. An absent or
 * unrecognized state reads as `visible` (the spec's default).
 */
export type SheetState = "visible" | "hidden" | "veryHidden";

export interface SheetInfo {
	/** Sheet name as shown on Excel's tab. */
	readonly name: string;
	/** Workbook-relative part path, resolved via the relationship graph. */
	readonly path: string;
	/** false for hidden or very-hidden sheets. Kept alongside {@link state} (which it derives from). */
	readonly visible: boolean;
	/** The tab's visibility state (F4.6). `visible` is `state === "visible"`. */
	readonly state: SheetState;
}

export interface Comment {
	/** The cell the comment is anchored to, e.g. "B2". */
	readonly ref: string;
	/** Comment author, resolved from the authors table. Absent when it can't be resolved. */
	readonly author?: string;
	/** The comment's plain text — rich-text runs concatenated, formatting dropped. */
	readonly text: string;
}

export interface Hyperlink {
	/** The cell or range the link covers, e.g. "A1" or "B1:C2". */
	readonly ref: string;
	/**
	 * External destination (a URL, `mailto:`, or `file:` target) resolved through the
	 * worksheet's relationships. Absent for a purely in-workbook link.
	 */
	readonly target?: string;
	/** In-workbook destination, e.g. "'Sheet2'!B5". Absent for a purely external link. */
	readonly location?: string;
	/** Hover text the producer attached to the link, if any. */
	readonly tooltip?: string;
	/** Display-text override for the link, if any. */
	readonly display?: string;
}

// ── Sheet geometry (F4.5) ──────────────────────────────────────────────────────────────────────
// One shared model, like styles: what the reader's accessors return IS what the writer accepts.

/** Width/visibility for a 1-based column range (`min`–`max` inclusive), from `<cols>`. */
export interface ColumnProps {
	readonly min: number;
	readonly max: number;
	/** Column width in characters of the default font (Excel's unit), 0 < width ≤ 255. */
	readonly width?: number;
	readonly hidden?: boolean;
}

/** Height/visibility of one row, from `<row ht hidden>`. */
export interface RowProps {
	/** Row height in points, 0 < height ≤ 409.5 (Excel's ceiling). */
	readonly height?: number;
	readonly hidden?: boolean;
}

/**
 * A frozen pane: the top `rows` rows and/or leftmost `cols` columns stay visible while the rest
 * scrolls. Split (non-frozen) panes are not modelled and read as no freeze.
 */
export interface FreezePane {
	readonly rows?: number;
	readonly cols?: number;
}

// ── Images (M6) ──────────────────────────────────────────────────────────────────────────────
// One shared model: what `Worksheet.images()` returns IS what the writer accepts (F6.3). Anchors
// are kept RAW like colors — cell col/row plus EMU offsets/extents verbatim, never converted to
// pixels (914 400 EMU/inch; ≈9 525 EMU/px @96 dpi) — because only the raw form round-trips exactly.

/** A drawing anchor point: a 1-based cell plus an EMU offset into that cell (0 at the cell edge). */
export interface AnchorPoint {
	/** 1-based column of the anchored cell (OOXML stores it 0-based; converted on read). */
	readonly col: number;
	/** 1-based row of the anchored cell. */
	readonly row: number;
	/** Horizontal offset into the cell, in EMU. Defaults to 0 when omitted. */
	readonly colOff?: number;
	/** Vertical offset into the cell, in EMU. */
	readonly rowOff?: number;
}

/**
 * How a picture is anchored. A `to` point (and no `ext`) is a two-cell anchor — the picture spans
 * `from`→`to` and resizes with the cells. An `ext` (and no `to`) is a one-cell anchor — pinned at
 * `from` with a fixed EMU size `{cx, cy}`. `editAs` is the producer's move/size behaviour when
 * present. Absolute-anchored pictures are not modelled (skipped on read).
 */
export interface ImageAnchor {
	readonly from: AnchorPoint;
	readonly to?: AnchorPoint;
	readonly ext?: { readonly cx: number; readonly cy: number };
	readonly editAs?: "twoCell" | "oneCell" | "absolute";
}

/**
 * A picture on a worksheet. `bytes` is the raw, undecoded image payload; `mime` is its media type
 * (`image/png`, `image/jpeg`, `image/gif`, …). `name` is the producer's picture name when present.
 * Pictures sharing one media part share one `bytes` buffer on read.
 */
export interface SheetImage {
	readonly anchor: ImageAnchor;
	readonly bytes: Uint8Array;
	readonly mime: string;
	readonly name?: string;
}

// ── Styles (M4) ────────────────────────────────────────────────────────────────────────────────
// One shared style model: what `Worksheet.style(ref)` returns IS what the writer accepts, so the
// read→modify→write bridge carries styles as a structural pass-through.

/**
 * A color as OOXML stores it — kept RAW, never resolved. `rgb` is ARGB hex (e.g. `"FFFF0000"`);
 * `theme` indexes the workbook theme's color scheme with an optional `tint` (−1…1); `indexed` is
 * a legacy palette index; `auto` lets the consumer pick (usually black). Theme colors are NOT
 * resolved to rgb on read: resolution needs a theme1.xml parser and is lossy on rewrite (a
 * theme-aware consumer could no longer re-tint) — the raw form is what round-trips faithfully,
 * and it is exactly what openpyxl stores too.
 */
export type Color =
	| { readonly rgb: string }
	| { readonly theme: number; readonly tint?: number }
	| { readonly indexed: number }
	| { readonly auto: true };

/**
 * Underline style. The exotic accounting variants (`singleAccounting`/`doubleAccounting`)
 * degrade to no underline on read and are rejected on write (deferred, documented).
 */
export type UnderlineStyle = "single" | "double";

export interface FontStyle {
	readonly name?: string;
	/** Font size in points. */
	readonly size?: number;
	readonly bold?: boolean;
	readonly italic?: boolean;
	readonly underline?: UnderlineStyle;
	readonly strike?: boolean;
	readonly color?: Color;
}

/** Fill pattern kinds (ECMA-376 §18.18.55). `gray125` is the workbook-reserved fill 1. */
export type PatternType =
	| "none"
	| "solid"
	| "mediumGray"
	| "darkGray"
	| "lightGray"
	| "darkHorizontal"
	| "darkVertical"
	| "darkDown"
	| "darkUp"
	| "darkGrid"
	| "darkTrellis"
	| "lightHorizontal"
	| "lightVertical"
	| "lightDown"
	| "lightUp"
	| "lightGrid"
	| "lightTrellis"
	| "gray125"
	| "gray0625";

/**
 * A pattern fill. For the everyday solid fill, the visible color is `fgColor` (OOXML's rule —
 * `bgColor` shows only through pattern gaps). Gradient fills are not modelled (deferred): a
 * gradient-filled cell reads as having no fill.
 */
export interface FillStyle {
	readonly patternType: PatternType;
	readonly fgColor?: Color;
	readonly bgColor?: Color;
}

/** Border line styles (ECMA-376 §18.18.3). An edge with no style is simply absent. */
export type BorderLineStyle =
	| "thin"
	| "medium"
	| "thick"
	| "dashed"
	| "dotted"
	| "double"
	| "hair"
	| "mediumDashed"
	| "dashDot"
	| "mediumDashDot"
	| "dashDotDot"
	| "mediumDashDotDot"
	| "slantDashDot";

export interface BorderEdge {
	readonly style: BorderLineStyle;
	readonly color?: Color;
}

/** Per-edge borders. Diagonal borders are not modelled (deferred). */
export interface BorderStyle {
	readonly top?: BorderEdge;
	readonly right?: BorderEdge;
	readonly bottom?: BorderEdge;
	readonly left?: BorderEdge;
}

export type HorizontalAlignment =
	| "left"
	| "center"
	| "right"
	| "justify"
	| "fill"
	| "centerContinuous"
	| "distributed";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distributed";

export interface Alignment {
	readonly horizontal?: HorizontalAlignment;
	readonly vertical?: VerticalAlignment;
	readonly wrapText?: boolean;
	readonly shrinkToFit?: boolean;
	/** Indent level (whole units of about 3 spaces), 0–250. */
	readonly indent?: number;
	/**
	 * Text rotation in degrees, 0–180 (91–180 mean 1–90° downward, per the spec). The legacy
	 * marker 255 ("vertical stacked") is not modelled and degrades to no rotation.
	 */
	readonly textRotation?: number;
}

/**
 * The resolved style of one cell. Every component is optional; a cell whose effective format is
 * the workbook default resolves to no style at all (`Worksheet.style(ref)` returns `undefined`).
 * `numberFormat` is always the format CODE string (e.g. `"yyyy-mm-dd"`, `"0.00%"`) — ids are a
 * file-internal detail and never appear in the API.
 */
export interface CellStyle {
	readonly numberFormat?: string;
	readonly font?: FontStyle;
	readonly fill?: FillStyle;
	readonly border?: BorderStyle;
	readonly alignment?: Alignment;
}

// ── The reader's worksheet surface (multi-format seam, M7) ───────────────────────────────────────
// `Worksheet` is a structural INTERFACE, not a class, so every format's reader can return the SAME
// public shape: the xlsx reader's `XlsxWorksheet` and the ODS reader's `OdsWorksheet` both implement
// it (F7.1). A format that can't express an accessor DEGRADES — returns `[]`/`undefined`/an empty
// map — never throws. `Workbook.sheet(name)` returns this type.

export interface Worksheet {
	/** Sheet name as shown on the tab. */
	readonly name: string;
	/** Format-native part locator (xlsx: `xl/worksheets/sheet1.xml`; ods: a synthetic id). */
	readonly path: string;
	/** false for hidden or very-hidden sheets. */
	readonly visible: boolean;
	/** The tab's visibility state: `"visible"`, `"hidden"`, or `"veryHidden"`. */
	readonly state: SheetState;
	/** Merged-cell ranges in A1 notation, in document order. Empty when none. */
	readonly mergedCells: readonly string[];
	/** Hyperlinks declared on this sheet, in document order. Empty when none. */
	readonly hyperlinks: readonly Hyperlink[];
	/** The sheet's declared used range in A1 notation, or `undefined` when absent. */
	readonly dimension: string | undefined;
	/** The comments anchored to cells on this sheet. Empty when none (or unsupported by the format). */
	readonly comments: readonly Comment[];
	/** Column width/visibility declarations, in document order. Empty when none (or unsupported). */
	readonly columns: readonly ColumnProps[];
	/** Per-row height/visibility, keyed by 1-based row index. Empty when none (or unsupported). */
	readonly rowProperties: ReadonlyMap<number, RowProps>;
	/** The sheet's frozen pane, or `undefined` when nothing is frozen (or unsupported). */
	readonly freeze: FreezePane | undefined;
	/** The number-format code applied at `ref`, or `undefined` (also when the format has no styles). */
	numberFormat(ref: string): string | undefined;
	/** The resolved style at `ref`, or `undefined` (also when the format has no styles). */
	style(ref: string): CellStyle | undefined;
	/** The formula text at `ref`, or `undefined` (also when the format doesn't carry formula text). */
	formula(ref: string): string | undefined;
	/** The pictures on this sheet, in order. Empty when none (or unsupported by the format). */
	images(): Promise<readonly SheetImage[]>;
	/** The cell at an A1 reference. Absent cells read as `empty`. */
	cell(ref: string): Cell;
	/** Stream the populated rows in document order. Sparse: empty rows/cells are absent. */
	rows(): AsyncGenerator<Row>;
}
