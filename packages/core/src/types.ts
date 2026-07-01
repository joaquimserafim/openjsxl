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
