import { XlsxError } from '../errors'
import { formatRef } from '../ooxml/a1'
import { dateToSerial } from '../ooxml/dates'
import type { CellValue } from './types'
import { escapeText, isXmlSafe, preserveAttr } from './xml'

// Serialize one sheet's rows into worksheet XML (`xl/worksheets/sheetN.xml`). The element order the
// schema requires here is <dimension> then <sheetData>; within <sheetData>, rows ascend by index and
// cells ascend by column — which is exactly the order we walk the input arrays.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

// The cellXfs index the writer's styles.xml reserves for a date (numFmtId 14). Index 0 is General.
// A date cell carries s="1"; the reader resolves that style to a date format and types it back as a
// `date` cell, closing the round-trip.
export const DATE_STYLE_INDEX = 1

interface RenderedCell {
	readonly xml: string
	readonly isDate: boolean
}

// A finite JS number as it appears in <v>. String() gives the shortest decimal that parses back to
// the same double via the reader's Number() — so the value round-trips exactly. Non-finite values
// (NaN, ±Infinity) have no .xlsx representation and are rejected before this.
function numberToXml(n: number): string {
	return String(n)
}

// Render a single cell, or `undefined` for an empty one (null/undefined value) omitted from the row.
// Throws `invalid-input` for a value that cannot be represented: a non-finite number, an invalid
// Date, or a type outside the CellValue union (JS callers can pass anything). `date1904` selects the
// serial epoch so it matches the workbook's declared <workbookPr date1904>.
function renderCell(
	col: number,
	row: number,
	value: CellValue,
	date1904: boolean,
): RenderedCell | undefined {
	if (value === null || value === undefined) return undefined
	const ref = formatRef({ col, row })

	if (typeof value === 'string') {
		// A forbidden control character or lone surrogate would make the part not well-formed (or be
		// silently mangled to U+FFFD by TextEncoder) — reject rather than emit a broken/lossy file.
		if (!isXmlSafe(value)) {
			throw new XlsxError(
				'invalid-input',
				`cell ${ref}: string contains a character not allowed in XML (a control character or lone surrogate)`,
			)
		}
		return {
			xml: `<c r="${ref}" t="inlineStr"><is><t${preserveAttr(value)}>${escapeText(value)}</t></is></c>`,
			isDate: false,
		}
	}
	if (typeof value === 'boolean') {
		return { xml: `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`, isDate: false }
	}
	if (typeof value === 'number') {
		if (!Number.isFinite(value)) {
			throw new XlsxError('invalid-input', `cell ${ref}: ${value} is not a finite number`)
		}
		return { xml: `<c r="${ref}"><v>${numberToXml(value)}</v></c>`, isDate: false }
	}
	if (value instanceof Date) {
		const serial = dateToSerial(value, date1904)
		if (!Number.isFinite(serial)) {
			throw new XlsxError('invalid-input', `cell ${ref}: invalid Date`)
		}
		return {
			xml: `<c r="${ref}" s="${DATE_STYLE_INDEX}"><v>${numberToXml(serial)}</v></c>`,
			isDate: true,
		}
	}
	throw new XlsxError('invalid-input', `cell ${ref}: unsupported cell value type`)
}

export interface WorksheetResult {
	readonly xml: string
	/** True if any cell is a date — the workbook emits styles.xml only when some sheet needs it. */
	readonly usesDate: boolean
}

/** Build the worksheet XML for one sheet. */
export function worksheetXml(
	rows: readonly (readonly CellValue[])[],
	date1904: boolean,
): WorksheetResult {
	let usesDate = false
	// 0 means "unset" — no populated cell seen yet (columns/rows are 1-based, so 0 is a safe sentinel).
	let minRow = 0
	let maxRow = 0
	let minCol = 0
	let maxCol = 0
	const rowXmls: string[] = []

	for (let r = 0; r < rows.length; r++) {
		const cells = rows[r]
		// A missing row (array hole / undefined) is an empty row — skip it. Anything else that isn't
		// an array (a string, a number, null, an object) would otherwise be iterated as if it were a
		// row: a string "abc" would explode into three character cells. Reject it instead of silently
		// mangling the data — this also turns a null row into a clean error rather than a TypeError.
		if (cells === undefined) continue
		if (!Array.isArray(cells)) {
			throw new XlsxError(
				'invalid-input',
				`sheet row ${r + 1}: a row must be an array of cell values`,
			)
		}
		if (cells.length === 0) continue
		const rowNum = r + 1
		const cellXmls: string[] = []
		for (let c = 0; c < cells.length; c++) {
			const colNum = c + 1
			const rendered = renderCell(colNum, rowNum, cells[c], date1904)
			if (rendered === undefined) continue
			if (rendered.isDate) usesDate = true
			if (minRow === 0 || rowNum < minRow) minRow = rowNum
			if (rowNum > maxRow) maxRow = rowNum
			if (minCol === 0 || colNum < minCol) minCol = colNum
			if (colNum > maxCol) maxCol = colNum
			cellXmls.push(rendered.xml)
		}
		if (cellXmls.length > 0) rowXmls.push(`<row r="${rowNum}">${cellXmls.join('')}</row>`)
	}

	// Bounding box of the populated cells, in A1 notation. An entirely empty sheet is "A1" (Excel's
	// convention); a single cell collapses to that one ref rather than a degenerate "X:X" range.
	const dimension =
		minRow === 0
			? 'A1'
			: minRow === maxRow && minCol === maxCol
				? formatRef({ col: minCol, row: minRow })
				: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`

	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"><dimension ref="${dimension}"/><sheetData>${rowXmls.join(
		'',
	)}</sheetData></worksheet>`
	return { xml, usesDate }
}
