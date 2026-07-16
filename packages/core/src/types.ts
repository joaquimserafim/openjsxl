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

// ── Tables (F9.1) ────────────────────────────────────────────────────────────────────────────
// One shared model: what `Worksheet.tables` returns IS what `SheetInput.tables` accepts, so a table
// round-trips through the bridge as a structural pass-through. The numeric table `id` and any raw
// dxf indexes are internal (auto-assigned / interned on write), never surfaced.

/** A single column of a {@link TableInfo}. */
export interface TableColumn {
	/** The column's name — must equal its header-row cell text (Excel repair-prompts otherwise). */
	readonly name: string;
	/** A literal label shown in the totals row, if the producer set one (carried verbatim). */
	readonly totalsRowLabel?: string;
	/** A built-in totals-row function (e.g. `"sum"`, `"count"`), if any (carried verbatim). */
	readonly totalsRowFunction?: string;
	/** A custom totals-row formula, if any — element text, carried verbatim, never evaluated. */
	readonly totalsRowFormula?: string;
	/** A calculated-column formula, if any — element text, carried verbatim, never evaluated. */
	readonly calculatedColumnFormula?: string;
	/**
	 * The header-cell highlight for this column (`headerRowDxfId`), resolved to an inline
	 * {@link DxfStyle} (F9.3). The numeric dxf index is never surfaced. Absent when the producer set none.
	 */
	readonly headerRowStyle?: DxfStyle;
	/** The data-cells highlight for this column (`dataDxfId`), resolved to an inline {@link DxfStyle}. */
	readonly dataStyle?: DxfStyle;
	/** The totals-cell highlight for this column (`totalsRowDxfId`), resolved to an inline {@link DxfStyle}. */
	readonly totalsRowStyle?: DxfStyle;
}

/** The built-in table-style banding for a {@link TableInfo} (`<tableStyleInfo>`). */
export interface TableStyleInfo {
	/** A built-in table-style name, e.g. `"TableStyleMedium9"`. */
	readonly name?: string;
	readonly showFirstColumn?: boolean;
	readonly showLastColumn?: boolean;
	readonly showRowStripes?: boolean;
	readonly showColumnStripes?: boolean;
}

/**
 * A defined table (`xl/tables/tableN.xml`) — a named, structured range with a header row. The reader
 * surfaces the producer's values; the writer derives column names from the header row and auto-assigns
 * the numeric id. Table-level dxf highlight indexes are dropped in F9.1 (they return as inline styles
 * once conditional formatting / dxfs land).
 */
export interface TableInfo {
	/** The table's display name — a workbook-unique identifier (no spaces, not a cell reference). */
	readonly name: string;
	/** The table's range in A1 notation, e.g. `"A1:C5"` (includes the header and any totals row). */
	readonly ref: string;
	/** The columns, left to right; `columns.length` equals the width of `ref`. */
	readonly columns: readonly TableColumn[];
	/** Whether the table has a header row (default true; `headerRowCount="0"` turns it off). */
	readonly headerRow: boolean;
	/** Whether the table shows a totals row (default false). */
	readonly totalsRow: boolean;
	/** The built-in style banding, or `undefined` when the producer set none. */
	readonly style?: TableStyleInfo;
	/** Table-wide header-row highlight (`headerRowDxfId`), resolved to an inline {@link DxfStyle} (F9.3). */
	readonly headerRowStyle?: DxfStyle;
	/** Table-wide data highlight (`dataDxfId`), resolved to an inline {@link DxfStyle}. */
	readonly dataStyle?: DxfStyle;
	/** Table-wide totals-row highlight (`totalsRowDxfId`), resolved to an inline {@link DxfStyle}. */
	readonly totalsRowStyle?: DxfStyle;
}

// ── Data validation (F9.2) ─────────────────────────────────────────────────────────────────────
// One shared model: what `Worksheet.dataValidations` returns IS what `SheetInput.dataValidations`
// accepts, so a rule crosses the bridge as a structural pass-through. Formula operands are carried
// as verbatim TEXT (never evaluated). Worksheet-level x14 validations (Excel 2010+ cross-sheet list
// sources) live in `<extLst>` and are skipped on read / never emitted — a named degradation.

/** A validation kind (`<dataValidation type>`), ECMA-376 ST_DataValidationType. */
export type DataValidationType =
	| "none"
	| "whole"
	| "decimal"
	| "list"
	| "date"
	| "time"
	| "textLength"
	| "custom";

/** The comparison a validation applies (`<dataValidation operator>`), ST_DataValidationOperator. */
export type DataValidationOperator =
	| "between"
	| "notBetween"
	| "equal"
	| "notEqual"
	| "lessThan"
	| "lessThanOrEqual"
	| "greaterThan"
	| "greaterThanOrEqual";

/** How Excel reacts to invalid input (`<dataValidation errorStyle>`), ST_DataValidationErrorStyle. */
export type DataValidationErrorStyle = "stop" | "warning" | "information";

/**
 * A data-validation rule (`<dataValidation>`) constraining one or more cell ranges. `formula1` and
 * `formula2` are the producer's operand TEXT, carried verbatim: for a `list` type `formula1` is
 * either a range (`$A$1:$A$3`, possibly cross-sheet) or an inline comma list in quotes (`"a,b,c"` —
 * the quotes ARE part of the text); for `whole`/`decimal`/`date`/`time`/`textLength` they are the
 * bound(s); for `custom` `formula1` is the expression.
 *
 * `showDropDown` uses the INTUITIVE sense: `true` means Excel shows the in-cell dropdown arrow for a
 * `list`. The file's `showDropDown` attribute is inverted (a `1` there SUPPRESSES the arrow), so the
 * reader and writer translate it — a file `1` reads as `false` here, and this `false` writes back as
 * `showDropDown="1"`.
 */
export interface DataValidation {
	/** The ranges this rule covers, one per `@sqref` token, in A1 notation (e.g. `["A1:A10", "C1"]`). */
	readonly sqref: readonly string[];
	/** The validation type; an absent/unrecognized `type` reads as `"none"`. */
	readonly type: DataValidationType;
	/** The comparison operator (for whole/decimal/date/time/textLength); absent for list/custom/none. */
	readonly operator?: DataValidationOperator;
	/** The first operand text in stored form (a leading `=` stripped) — a bound, list source, or expression. */
	readonly formula1?: string;
	/** The second operand text in stored form — only meaningful for `between`/`notBetween`. */
	readonly formula2?: string;
	/** Whether an empty cell passes validation (`allowBlank`). */
	readonly allowBlank?: boolean;
	/**
	 * Whether Excel shows the in-cell dropdown for a `list` validation — INTUITIVE sense (`true` =
	 * arrow shown). The file attribute is inverted; see the interface note. Absent means Excel's
	 * default (arrow shown).
	 */
	readonly showDropDown?: boolean;
	/** Whether the input-message popup is shown (`showInputMessage`). */
	readonly showInputMessage?: boolean;
	/** Whether the error alert is shown on invalid input (`showErrorMessage`). */
	readonly showErrorMessage?: boolean;
	/** The error-alert style; absent means Excel's default (`stop`). */
	readonly errorStyle?: DataValidationErrorStyle;
	/** The input-message title (≤32 characters). */
	readonly promptTitle?: string;
	/** The input-message body (≤255 characters). */
	readonly prompt?: string;
	/** The error-alert title (≤32 characters). */
	readonly errorTitle?: string;
	/** The error-alert body (≤255 characters). */
	readonly error?: string;
}

// ── Differential styles + conditional formatting (F9.3) ────────────────────────────────────────
// One shared model: what `Worksheet.conditionalFormatting` returns IS what `SheetInput
// .conditionalFormatting` accepts. A rule's highlight look is an INLINE {@link DxfStyle} — the file's
// numeric `dxfId` (an index into styles.xml's `<dxfs>`) is never public: the reader resolves it to an
// inline style, the writer interns inline styles back into one `<dxfs>` table and re-assigns ids.

/**
 * A differential FILL — kept RAW, never normalized (unlike {@link FillStyle}). For a solid
 * conditional-formatting highlight the visible color is `bgColor` (the exact INVERSE of a cell fill,
 * where it is `fgColor`), and `patternType` is usually absent — normalizing here would silently swap
 * every highlight color, so the components are carried exactly as written.
 */
export interface DxfFill {
	readonly patternType?: PatternType;
	readonly fgColor?: Color;
	readonly bgColor?: Color;
}

/**
 * A differential ("dxf") style — the partial format a conditional-formatting rule (or, later, a table
 * region) applies over the cells it matches. Only the components it OVERRIDES are present. It shares
 * the {@link FontStyle}/{@link BorderStyle}/{@link Alignment}/{@link Color} primitives with
 * {@link CellStyle} but keeps its fill RAW (see {@link DxfFill}).
 */
export interface DxfStyle {
	readonly numberFormat?: string;
	readonly font?: FontStyle;
	readonly fill?: DxfFill;
	readonly border?: BorderStyle;
	readonly alignment?: Alignment;
}

/**
 * A conditional-format value object (`<cfvo>`) — a threshold for a color scale, data bar, or icon
 * set. `type` is `num`/`percent`/`max`/`min`/`formula`/`percentile` (verbatim); `val` is the raw
 * threshold text (a number or a formula, carried verbatim); `gte` (data bar / icon set) marks a
 * `≥` rather than `>` boundary.
 */
export interface Cfvo {
	readonly type: string;
	readonly val?: string;
	readonly gte?: boolean;
}

/** The ST_CfType values whose rule applies a dxf highlight (as opposed to a scale/bar/icon set). */
export type CfHighlightType =
	| "cellIs"
	| "expression"
	| "top10"
	| "aboveAverage"
	| "uniqueValues"
	| "duplicateValues"
	| "containsText"
	| "notContainsText"
	| "beginsWith"
	| "endsWith"
	| "containsBlanks"
	| "notContainsBlanks"
	| "containsErrors"
	| "notContainsErrors"
	| "timePeriod";

interface CfRuleBase {
	/** Rule precedence — lower wins when rules overlap. Renumbered densely 1..n on write. */
	readonly priority: number;
	/** When true, no lower-priority rule is evaluated for a matched cell. */
	readonly stopIfTrue?: boolean;
}

/**
 * A highlight rule — one that paints matching cells with a {@link DxfStyle}. Discriminated by `type`.
 * `formulas` are the `<formula>` children verbatim (0–3); the declarative attrs (`operator`/`text`/
 * `timePeriod`/`rank`/…) and any generated formula BOTH pass through untouched — regenerating either
 * would desynchronize a dual-encoded rule.
 */
export interface CfHighlightRule extends CfRuleBase {
	readonly type: CfHighlightType;
	/** The look applied to matched cells; absent when the producer set no `dxfId`. */
	readonly dxf?: DxfStyle;
	readonly operator?: string;
	readonly text?: string;
	readonly timePeriod?: string;
	readonly rank?: number;
	readonly percent?: boolean;
	readonly bottom?: boolean;
	readonly aboveAverage?: boolean;
	readonly equalAverage?: boolean;
	readonly stdDev?: number;
	readonly formulas?: readonly string[];
}

/** A 2- or 3-stop color scale (`type: "colorScale"`). `cfvo` and `colors` are positional and equal-length. */
export interface CfColorScaleRule extends CfRuleBase {
	readonly type: "colorScale";
	readonly cfvo: readonly Cfvo[];
	readonly colors: readonly Color[];
}

/** A data bar (`type: "dataBar"`) — two cfvos (min/max) and the bar `color`. */
export interface CfDataBarRule extends CfRuleBase {
	readonly type: "dataBar";
	readonly cfvo: readonly Cfvo[];
	readonly color: Color;
}

/** An icon set (`type: "iconSet"`) — a built-in `iconSet` name and one cfvo per icon. */
export interface CfIconSetRule extends CfRuleBase {
	readonly type: "iconSet";
	readonly iconSet?: string;
	readonly cfvo: readonly Cfvo[];
}

/** A conditional-formatting rule — discriminated by `type`. */
export type ConditionalFormattingRule =
	| CfHighlightRule
	| CfColorScaleRule
	| CfDataBarRule
	| CfIconSetRule;

/**
 * One `<conditionalFormatting>` block: the ranges it covers (`sqref`, symbolic A1) and the rules that
 * apply over them, in document order.
 */
export interface ConditionalFormatting {
	readonly sqref: readonly string[];
	readonly rules: readonly ConditionalFormattingRule[];
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

/**
 * A sheet's autoFilter — the range that carries Excel's filter-dropdown arrows (F10.2). `ref` is an
 * A1 range like `"A1:C10"`. Only the range is modelled; per-column filter criteria and sort state are
 * not carried (a documented drop). Excel also records this range as a hidden `_xlnm._FilterDatabase`
 * defined name, which the reader folds into this field and the writer synthesizes from it — so a filter
 * has a single representation, never a duplicated defined name.
 */
export interface SheetAutoFilter {
	readonly ref: string;
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
	/** The defined tables on this sheet, in document order. Empty when none (or unsupported). */
	readonly tables: readonly TableInfo[];
	/** The data-validation rules on this sheet, in document order. Empty when none (or unsupported). */
	readonly dataValidations: readonly DataValidation[];
	/** The conditional-formatting blocks on this sheet, in document order. Empty when none (or unsupported). */
	readonly conditionalFormatting: readonly ConditionalFormatting[];
	/** The sheet's autoFilter range (filter dropdowns), or `undefined` when none (or unsupported). */
	readonly autoFilter: SheetAutoFilter | undefined;
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
