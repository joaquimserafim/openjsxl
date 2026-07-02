// Public cell model. A discriminated union, so narrowing on `type` also narrows
// `value`: `if (cell.type === 'date') { cell.value /* : Date */ }`.

export type CellType = 'empty' | 'string' | 'number' | 'boolean' | 'date' | 'error'

interface CellBase {
	/** A1 reference, e.g. "B2". */
	readonly ref: string
}

export type Cell =
	| (CellBase & { readonly type: 'empty'; readonly value: null })
	| (CellBase & { readonly type: 'string'; readonly value: string })
	| (CellBase & { readonly type: 'number'; readonly value: number })
	| (CellBase & { readonly type: 'boolean'; readonly value: boolean })
	| (CellBase & { readonly type: 'date'; readonly value: Date })
	| (CellBase & { readonly type: 'error'; readonly value: string })

export interface SheetInfo {
	/** Sheet name as shown on Excel's tab. */
	readonly name: string
	/** Workbook-relative part path, resolved via the relationship graph. */
	readonly path: string
	/** false for hidden or very-hidden sheets. */
	readonly visible: boolean
}

export interface Comment {
	/** The cell the comment is anchored to, e.g. "B2". */
	readonly ref: string
	/** Comment author, resolved from the authors table. Absent when it can't be resolved. */
	readonly author?: string
	/** The comment's plain text — rich-text runs concatenated, formatting dropped. */
	readonly text: string
}

export interface Hyperlink {
	/** The cell or range the link covers, e.g. "A1" or "B1:C2". */
	readonly ref: string
	/**
	 * External destination (a URL, `mailto:`, or `file:` target) resolved through the
	 * worksheet's relationships. Absent for a purely in-workbook link.
	 */
	readonly target?: string
	/** In-workbook destination, e.g. "'Sheet2'!B5". Absent for a purely external link. */
	readonly location?: string
	/** Hover text the producer attached to the link, if any. */
	readonly tooltip?: string
	/** Display-text override for the link, if any. */
	readonly display?: string
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
	| { readonly auto: true }

/**
 * Underline style. The exotic accounting variants (`singleAccounting`/`doubleAccounting`)
 * degrade to no underline on read and are rejected on write (deferred, documented).
 */
export type UnderlineStyle = 'single' | 'double'

export interface FontStyle {
	readonly name?: string
	/** Font size in points. */
	readonly size?: number
	readonly bold?: boolean
	readonly italic?: boolean
	readonly underline?: UnderlineStyle
	readonly strike?: boolean
	readonly color?: Color
}

/** Fill pattern kinds (ECMA-376 §18.18.55). `gray125` is the workbook-reserved fill 1. */
export type PatternType =
	| 'none'
	| 'solid'
	| 'mediumGray'
	| 'darkGray'
	| 'lightGray'
	| 'darkHorizontal'
	| 'darkVertical'
	| 'darkDown'
	| 'darkUp'
	| 'darkGrid'
	| 'darkTrellis'
	| 'lightHorizontal'
	| 'lightVertical'
	| 'lightDown'
	| 'lightUp'
	| 'lightGrid'
	| 'lightTrellis'
	| 'gray125'
	| 'gray0625'

/**
 * A pattern fill. For the everyday solid fill, the visible color is `fgColor` (OOXML's rule —
 * `bgColor` shows only through pattern gaps). Gradient fills are not modelled (deferred): a
 * gradient-filled cell reads as having no fill.
 */
export interface FillStyle {
	readonly patternType: PatternType
	readonly fgColor?: Color
	readonly bgColor?: Color
}

/** Border line styles (ECMA-376 §18.18.3). An edge with no style is simply absent. */
export type BorderLineStyle =
	| 'thin'
	| 'medium'
	| 'thick'
	| 'dashed'
	| 'dotted'
	| 'double'
	| 'hair'
	| 'mediumDashed'
	| 'dashDot'
	| 'mediumDashDot'
	| 'dashDotDot'
	| 'mediumDashDotDot'
	| 'slantDashDot'

export interface BorderEdge {
	readonly style: BorderLineStyle
	readonly color?: Color
}

/** Per-edge borders. Diagonal borders are not modelled (deferred). */
export interface BorderStyle {
	readonly top?: BorderEdge
	readonly right?: BorderEdge
	readonly bottom?: BorderEdge
	readonly left?: BorderEdge
}

export type HorizontalAlignment =
	| 'left'
	| 'center'
	| 'right'
	| 'justify'
	| 'fill'
	| 'centerContinuous'
	| 'distributed'

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed'

export interface Alignment {
	readonly horizontal?: HorizontalAlignment
	readonly vertical?: VerticalAlignment
	readonly wrapText?: boolean
	readonly shrinkToFit?: boolean
	/** Indent level (whole units of about 3 spaces), 0–250. */
	readonly indent?: number
	/**
	 * Text rotation in degrees, 0–180 (91–180 mean 1–90° downward, per the spec). The legacy
	 * marker 255 ("vertical stacked") is not modelled and degrades to no rotation.
	 */
	readonly textRotation?: number
}

/**
 * The resolved style of one cell. Every component is optional; a cell whose effective format is
 * the workbook default resolves to no style at all (`Worksheet.style(ref)` returns `undefined`).
 * `numberFormat` is always the format CODE string (e.g. `"yyyy-mm-dd"`, `"0.00%"`) — ids are a
 * file-internal detail and never appear in the API.
 */
export interface CellStyle {
	readonly numberFormat?: string
	readonly font?: FontStyle
	readonly fill?: FillStyle
	readonly border?: BorderStyle
	readonly alignment?: Alignment
}
