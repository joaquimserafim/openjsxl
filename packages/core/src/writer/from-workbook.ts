import { parseRef } from '../ooxml/a1'
import type { Workbook } from '../reader/workbook'
import type { Cell } from '../types'
import type { CellValue, SheetInput, WorkbookInput } from './types'

// Bridge the reader to the writer: turn an open Workbook into the plain-data input writeXlsx wants,
// so a file can be read, optionally tweaked, and written back. This closes the round trip (F3.3).
//
// Fidelity is scoped to values, types, and sheet names/order — the writer's supported set. What the
// writer can't represent is NOT carried across and is documented, not silently mangled:
//   - formulas               → only the cached value survives (its type/number)
//   - error cells            → written as their literal text (e.g. "#DIV/0!"), so they become strings
//   - merges, hyperlinks, comments, custom number formats, sheet visibility → dropped (the M4 surface)
// A source sheet name the writer rejects (>31 chars, forbidden characters, duplicate) will make the
// subsequent writeXlsx throw invalid-input — such a workbook isn't Excel-valid to re-emit as-is.

// The reader already exposes each cell as its JS value: null for empty, and string / number /
// boolean / Date for the typed kinds. The lone nuance is 'error', whose value is the error text —
// the writer has no error-cell form, so it round-trips as that string. Hence a plain pass-through,
// wrapped for a single documented place to change if first-class error/formula support ever lands.
function cellToValue(cell: Cell): CellValue {
	return cell.value
}

/**
 * Convert an open {@link Workbook} into {@link WorkbookInput} for {@link writeXlsx}. Each populated
 * cell is placed at its own A1 reference, preserving sheet names and tab order. Rows/columns are
 * left sparse (array holes) — the writer treats a hole as an empty cell — so a workbook with a few
 * far-apart cells does not materialize a dense grid.
 */
export async function workbookToInput(workbook: Workbook): Promise<WorkbookInput> {
	const sheets: SheetInput[] = []
	for (const info of workbook.sheets) {
		const worksheet = workbook.sheet(info.name)
		const rows: CellValue[][] = []
		for await (const row of worksheet.rows()) {
			for (const cell of row.cells) {
				// Place by the cell's own ref, not the row index — the two agree for well-formed
				// files, and the ref is authoritative for the column in either case.
				const { col, row: rowNum } = parseRef(cell.ref)
				let rowArr = rows[rowNum - 1]
				if (rowArr === undefined) {
					rowArr = []
					rows[rowNum - 1] = rowArr
				}
				rowArr[col - 1] = cellToValue(cell)
			}
		}
		sheets.push({ name: info.name, rows })
	}
	return { sheets }
}
