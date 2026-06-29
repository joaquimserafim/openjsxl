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
