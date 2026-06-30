// Public cell model. A discriminated union, so narrowing on `type` also narrows
// `value`: `if (cell.type === 'date') { cell.value /* : Date */ }`.

export type CellType = 'empty' | 'string' | 'number' | 'boolean' | 'date' | 'error'

interface CellBase {
	/** A1 reference, e.g. "B2". */
	ref: string
}

export type Cell =
	| (CellBase & { type: 'empty'; value: null })
	| (CellBase & { type: 'string'; value: string })
	| (CellBase & { type: 'number'; value: number })
	| (CellBase & { type: 'boolean'; value: boolean })
	| (CellBase & { type: 'date'; value: Date })
	| (CellBase & { type: 'error'; value: string })

export interface SheetInfo {
	/** Sheet name as shown on Excel's tab. */
	name: string
	/** Workbook-relative part path, resolved via the relationship graph. */
	path: string
	/** false for hidden or very-hidden sheets. */
	visible: boolean
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
