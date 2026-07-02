import { XlsxError } from "../errors"
import { formatRef, MAX_COL, MAX_ROW } from "../ooxml/a1"
import { dateToSerial } from "../ooxml/dates"
import type { StyleRegistry } from "./styles"
import type { CellInput, CellValue, StyledCell } from "./types"
import { escapeText, isXmlSafe, preserveAttr } from "./xml"

// Serialize one sheet's rows into worksheet XML (`xl/worksheets/sheetN.xml`). The element order the
// schema requires here is <dimension> then <sheetData>; within <sheetData>, rows ascend by index and
// cells ascend by column — which is exactly the order we walk the input arrays.
//
// Styles (F4.2): every cell resolves an xf index through the shared StyleRegistry — 0 (the default)
// omits the `s` attribute entirely, so bare-value input emits the exact pre-F4.2 bytes. A styled
// BLANK cell ({ value: null, style }) is real: it emits a valueless `<c r s/>` and counts toward
// the dimension, which is how a border or fill lands on an empty cell.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// A finite JS number as it appears in <v>. String() gives the shortest decimal that parses back to
// the same double via the reader's Number() — so the value round-trips exactly. Non-finite values
// (NaN, ±Infinity) have no .xlsx representation and are rejected before this.
function numberToXml(n: number): string {
	return String(n)
}

// Split a CellInput into its value and (optional) style. Discrimination is total: null/undefined
// are empty, Date instances are dates, primitives are bare values — any OTHER object must be a
// StyledCell. One without a `value` property is some stray object (the pre-F4.2 writer rejected
// every object; keeping that loudness catches typos like { val: 1 } or a nested array).
function splitInput(
	col: number,
	row: number,
	input: CellInput,
): { readonly value: CellValue; readonly styled: StyledCell | undefined } {
	if (input === null || input === undefined) return { value: input, styled: undefined }
	if (typeof input !== "object" || input instanceof Date) {
		return { value: input, styled: undefined }
	}
	const ref = formatRef({ col, row })
	if (Array.isArray(input)) {
		throw new XlsxError("invalid-input", `cell ${ref}: an array is not a cell value`)
	}
	const record = input as unknown as Record<string, unknown>
	for (const key of Object.keys(record)) {
		if (key !== "value" && key !== "style") {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: a styled cell allows only "value" and "style" (got "${key}")`,
			)
		}
	}
	if (!("value" in record)) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: an object cell must be { value, style? } — did you mean a StyledCell?`,
		)
	}
	const value = record.value
	// The inner value must be a plain CellValue — a nested { value } or any other object (except
	// Date) has no meaning and would silently mis-serialize.
	if (
		value !== null &&
		value !== undefined &&
		typeof value === "object" &&
		!(value instanceof Date)
	) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: a styled cell's value cannot be an object`,
		)
	}
	return { value: value as CellValue, styled: input as StyledCell }
}

// Render a single cell, or `undefined` for one that produces no output (an empty value with no
// effective style). Throws `invalid-input` for a value that cannot be represented: a non-finite
// number, an invalid Date, or a type outside the CellInput union (JS callers can pass anything).
// `date1904` selects the serial epoch so it matches the workbook's declared <workbookPr date1904>.
function renderCell(
	col: number,
	row: number,
	input: CellInput,
	date1904: boolean,
	styles: StyleRegistry,
): string | undefined {
	const { value, styled } = splitInput(col, row, input)
	const ref = formatRef({ col, row })

	// Resolve the xf index. Bare non-date values never touch the registry (zero overhead on the
	// unstyled path); a Date always does (it needs the date number format).
	let xf = 0
	if (value instanceof Date) {
		xf = styles.xfIndexFor(styled?.style, true, ref)
	} else if (styled?.style !== undefined) {
		xf = styles.xfIndexFor(styled.style, false, ref)
	}
	const sAttr = xf === 0 ? "" : ` s="${xf}"`

	if (value === null || value === undefined) {
		// A styled blank emits a valueless cell; an unstyled (or default-styled) empty is omitted.
		return xf === 0 ? undefined : `<c r="${ref}"${sAttr}/>`
	}
	if (typeof value === "string") {
		// A forbidden control character or lone surrogate would make the part not well-formed (or be
		// silently mangled to U+FFFD by TextEncoder) — reject rather than emit a broken/lossy file.
		if (!isXmlSafe(value)) {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: string contains a character not allowed in XML (a control character or lone surrogate)`,
			)
		}
		return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${preserveAttr(value)}>${escapeText(value)}</t></is></c>`
	}
	if (typeof value === "boolean") {
		return `<c r="${ref}"${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`
	}
	if (typeof value === "number") {
		if (!Number.isFinite(value)) {
			throw new XlsxError("invalid-input", `cell ${ref}: ${value} is not a finite number`)
		}
		return `<c r="${ref}"${sAttr}><v>${numberToXml(value)}</v></c>`
	}
	if (value instanceof Date) {
		const serial = dateToSerial(value, date1904)
		if (!Number.isFinite(serial)) {
			throw new XlsxError("invalid-input", `cell ${ref}: invalid Date`)
		}
		return `<c r="${ref}"${sAttr}><v>${numberToXml(serial)}</v></c>`
	}
	throw new XlsxError("invalid-input", `cell ${ref}: unsupported cell value type`)
}

export interface WorksheetResult {
	readonly xml: string
}

/** Build the worksheet XML for one sheet, interning cell styles into the shared registry. */
export function worksheetXml(
	rows: readonly (readonly CellInput[])[],
	date1904: boolean,
	styles: StyleRegistry,
): WorksheetResult {
	// 0 means "unset" — no populated cell seen yet (columns/rows are 1-based, so 0 is a safe sentinel).
	let minRow = 0
	let maxRow = 0
	let minCol = 0
	let maxCol = 0
	const rowXmls: string[] = []

	// A workbook can't outgrow Excel's grid: refs past XFD1048576 make Excel refuse the file, and
	// (mechanically) `rows.length` drives this loop — an absurd length would spin for hours.
	if (rows.length > MAX_ROW) {
		throw new XlsxError("invalid-input", `a sheet cannot have more than ${MAX_ROW} rows`)
	}

	for (let r = 0; r < rows.length; r++) {
		const cells = rows[r]
		// A missing row (array hole / undefined) is an empty row — skip it. Anything else that isn't
		// an array (a string, a number, null, an object) would otherwise be iterated as if it were a
		// row: a string "abc" would explode into three character cells. Reject it instead of silently
		// mangling the data — this also turns a null row into a clean error rather than a TypeError.
		if (cells === undefined) continue
		if (!Array.isArray(cells)) {
			throw new XlsxError(
				"invalid-input",
				`sheet row ${r + 1}: a row must be an array of cell values`,
			)
		}
		if (cells.length === 0) continue
		if (cells.length > MAX_COL) {
			throw new XlsxError(
				"invalid-input",
				`sheet row ${r + 1}: a row cannot have more than ${MAX_COL} cells`,
			)
		}
		const rowNum = r + 1
		const cellXmls: string[] = []
		for (let c = 0; c < cells.length; c++) {
			const colNum = c + 1
			const rendered = renderCell(colNum, rowNum, cells[c], date1904, styles)
			if (rendered === undefined) continue
			if (minRow === 0 || rowNum < minRow) minRow = rowNum
			if (rowNum > maxRow) maxRow = rowNum
			if (minCol === 0 || colNum < minCol) minCol = colNum
			if (colNum > maxCol) maxCol = colNum
			cellXmls.push(rendered)
		}
		if (cellXmls.length > 0) rowXmls.push(`<row r="${rowNum}">${cellXmls.join("")}</row>`)
	}

	// Bounding box of the populated cells, in A1 notation. An entirely empty sheet is "A1" (Excel's
	// convention); a single cell collapses to that one ref rather than a degenerate "X:X" range.
	const dimension =
		minRow === 0
			? "A1"
			: minRow === maxRow && minCol === maxCol
				? formatRef({ col: minCol, row: minRow })
				: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`

	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"><dimension ref="${dimension}"/><sheetData>${rowXmls.join(
		"",
	)}</sheetData></worksheet>`
	return { xml }
}
