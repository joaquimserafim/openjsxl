import { XlsxError } from "../errors"
import { type CellRef, formatRef, MAX_COL, MAX_ROW, parseRef } from "../ooxml/a1"
import type { Workbook, Worksheet } from "../reader/workbook"
import type {
	Cell,
	ColumnProps,
	Comment,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetState,
} from "../types"
import type { CellInput, SheetInput, WorkbookInput } from "./types"

// Bridge the reader to the writer: turn an open Workbook into the plain-data input writeXlsx wants,
// so a file can be read, optionally tweaked, and written back. This closes the round trip (F3.3);
// since F4.4 it carries STYLES too — the CellStyle that sheet.style(ref) resolves is exactly the
// shape writeXlsx accepts, so styles cross the bridge as a structural pass-through.
//
// Fidelity is scoped to values, types, sheet names/order, and cell styles (number format, font,
// fill, border, alignment). What the writer can't represent is NOT carried across and is
// documented, not silently mangled:
//   - formulas               → only the cached value survives (its type/number)
//   - error cells            → written as their literal text (e.g. "#DIV/0!"), so they become strings
//   (column widths, row heights, hidden flags, frozen panes — F4.5 — merges, hyperlinks, sheet
//   visibility — F4.6 — and cell comments — F5.2 — DO carry across)
//   - locale-only number formats (ids 23–36, 50–58) → no portable code exists; the format flattens
//     (a date keeps its value and date-ness via the implicit format)
// Three documented FLATTENINGS (values stay exact; the file's internal spelling normalizes):
//   - a tolerated NON-CANONICAL cell ref spelling (e.g. lowercase "a1") re-emits canonically
//     ("A1") — same grid slot, same value/type/style; A1 notation is case-insensitive and the
//     writer's input is positional, so canonical is the only possible emission.
//   - row/column DEFAULT styles resolve into per-cell styles (the effective style each cell already
//     shows through style(ref)) — the rewritten file styles cells directly instead of via defaults.
//   - files authored under a CUSTOM THEME keep their raw {theme, tint} indexes, but the rewritten
//     package embeds the standard Office theme, so those indexes re-render against DEFAULT theme
//     colors. rgb/indexed colors are unaffected.
// A source sheet name the writer rejects (>31 chars, forbidden characters, duplicate) will make the
// subsequent writeXlsx throw invalid-input — such a workbook isn't Excel-valid to re-emit as-is.
// Likewise a cell whose `r` attribute is unaddressable garbage (the tolerant reader keeps it,
// faithful to the document, e.g. a 300-letter column ref): it has no grid position to re-write to,
// so the bridge throws a TYPED invalid-input instead of silently dropping it or crashing bare.

// The reader already exposes each cell as its JS value: null for empty, and string / number /
// boolean / Date for the typed kinds. The lone nuance is 'error', whose value is the error text —
// the writer has no error-cell form, so it round-trips as that string. A cell with a resolved
// style wraps into the writer's { value, style } form — including a styled EMPTY cell (<c s/>),
// which is how a border or fill on a blank cell survives the trip. style(ref) returns one cached
// object per distinct format, so the wrapping adds O(distinct formats) objects, not O(cells).
function cellToInput(worksheet: Worksheet, cell: Cell): CellInput {
	const style = worksheet.style(cell.ref)
	return style === undefined ? cell.value : { value: cell.value, style }
}

/**
 * Convert an open {@link Workbook} into {@link WorkbookInput} for {@link writeXlsx}. Each populated
 * cell is placed at its own A1 reference with its resolved style (if any), preserving sheet names
 * and tab order. Rows/columns are left sparse (array holes) — the writer treats a hole as an empty
 * cell — so a workbook with a few far-apart cells does not materialize a dense grid. An unstyled
 * workbook yields bare values only: the exact pre-styles WorkbookInput, and the exact same bytes
 * on rewrite.
 */
export async function workbookToInput(workbook: Workbook): Promise<WorkbookInput> {
	const sheets: SheetInput[] = []
	for (const info of workbook.sheets) {
		const worksheet = workbook.sheet(info.name)
		const rows: CellInput[][] = []
		// Verbatim source ref per occupied grid slot (canonical ref → the ref string that filled
		// it). The reader's cell identity is the VERBATIM ref: "A1" and "a1" are two different
		// cells to cell(), yet parse to one grid slot — silently letting one overwrite the other
		// would vanish a value with no error. Same-spelling duplicates stay last-wins, which is
		// exactly how the reader's own cell() map resolves them, so the two sides agree.
		const occupied = new Map<string, string>()
		for await (const row of worksheet.rows()) {
			for (const cell of row.cells) {
				// Place by the cell's own ref, not the row index — the two agree for well-formed
				// files, and the ref is authoritative for the column in either case.
				let placed: CellRef | undefined
				try {
					placed = parseRef(cell.ref)
				} catch {
					placed = undefined
				}
				// A ref that doesn't parse OR lies outside Excel's grid has no writable position.
				// The grid cap also matters mechanically: `rows` is indexed by row number, so a
				// hostile row like 1e14 would otherwise become an array LENGTH the writer iterates.
				if (placed === undefined || placed.row > MAX_ROW || placed.col > MAX_COL) {
					const shown = cell.ref.length > 24 ? `${cell.ref.slice(0, 24)}…` : cell.ref
					throw new XlsxError(
						"invalid-input",
						`sheet "${info.name}": cell reference "${shown}" has no writable grid position`,
					)
				}
				const canonical = formatRef(placed)
				const prior = occupied.get(canonical)
				if (prior !== undefined && prior !== cell.ref) {
					throw new XlsxError(
						"invalid-input",
						`sheet "${info.name}": cells "${prior}" and "${cell.ref}" are distinct to the reader but occupy one grid position (${canonical})`,
					)
				}
				occupied.set(canonical, cell.ref)
				const { col, row: rowNum } = placed
				let rowArr = rows[rowNum - 1]
				if (rowArr === undefined) {
					rowArr = []
					rows[rowNum - 1] = rowArr
				}
				rowArr[col - 1] = cellToInput(worksheet, cell)
			}
		}
		// Geometry (F4.5) and structural metadata (F4.6): carried only when present, so a workbook
		// using neither produces the exact same WorkbookInput — and the exact same bytes — as before.
		const sheet: {
			name: string
			rows: CellInput[][]
			columns?: readonly ColumnProps[]
			rowProperties?: Readonly<Record<number, RowProps>>
			freeze?: FreezePane
			merges?: readonly string[]
			hyperlinks?: readonly Hyperlink[]
			state?: SheetState
			comments?: readonly Comment[]
		} = { name: info.name, rows }
		const columns = worksheet.columns
		if (columns.length > 0) sheet.columns = columns
		const rowProperties = worksheet.rowProperties
		if (rowProperties.size > 0) sheet.rowProperties = Object.fromEntries(rowProperties)
		const freeze = worksheet.freeze
		if (freeze !== undefined) sheet.freeze = freeze
		const merges = worksheet.mergedCells
		if (merges.length > 0) sheet.merges = merges
		const hyperlinks = worksheet.hyperlinks
		if (hyperlinks.length > 0) sheet.hyperlinks = hyperlinks
		if (info.state !== "visible") sheet.state = info.state
		const comments = worksheet.comments
		if (comments.length > 0) sheet.comments = comments
		sheets.push(sheet)
	}
	return { sheets }
}
