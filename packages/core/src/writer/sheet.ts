import { XlsxError } from "../errors"
import { formatRef, MAX_COL, MAX_COL_WIDTH, MAX_ROW, MAX_ROW_HEIGHT } from "../ooxml/a1"
import { dateToSerial } from "../ooxml/dates"
import type { ColumnProps, FreezePane, RowProps } from "../types"
import type { StyleRegistry } from "./styles"
import type { CellInput, CellValue, SheetInput, StyledCell } from "./types"
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

// ── Sheet geometry (F4.5): validation + emission ───────────────────────────────────────────────
// Same philosophy as styles: strict validation naming the sheet, `false`/empty normalize away,
// and the accepted bounds are exactly what the reader's geometry accessors can produce — so the
// bridge's geometry always writes.

function sheetInvalid(sheetName: string, message: string): never {
	throw new XlsxError("invalid-input", `sheet "${sheetName}": ${message}`)
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false
	const proto = Object.getPrototypeOf(value)
	return proto === null || proto === Object.prototype
}

function checkGeoKeys(
	sheetName: string,
	what: string,
	obj: Record<string, unknown>,
	allowed: readonly string[],
): void {
	for (const key of Object.keys(obj)) {
		if (!allowed.includes(key))
			sheetInvalid(sheetName, `${what} has an unknown property "${key}"`)
	}
}

// <cols> — one <col> per entry that survives normalization ({hidden: false} alone melts away).
function colsXml(sheetName: string, columns: readonly ColumnProps[] | undefined): string {
	if (columns === undefined) return ""
	if (!Array.isArray(columns)) sheetInvalid(sheetName, "columns must be an array")
	const entries: string[] = []
	for (let i = 0; i < columns.length; i++) {
		const raw = columns[i] as unknown
		const what = `columns[${i}]`
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`)
		checkGeoKeys(sheetName, what, raw, ["min", "max", "width", "hidden"])
		const min = raw.min
		const max = raw.max
		if (
			typeof min !== "number" ||
			!Number.isInteger(min) ||
			typeof max !== "number" ||
			!Number.isInteger(max) ||
			min < 1 ||
			max < min ||
			max > MAX_COL
		) {
			sheetInvalid(
				sheetName,
				`${what} needs integer 1-based min ≤ max within Excel's ${MAX_COL} columns`,
			)
		}
		let attrs = ` min="${min}" max="${max}"`
		const width = raw.width
		if (width !== undefined) {
			if (
				typeof width !== "number" ||
				!Number.isFinite(width) ||
				width <= 0 ||
				width > MAX_COL_WIDTH
			) {
				sheetInvalid(sheetName, `${what}.width must be a number in (0, ${MAX_COL_WIDTH}]`)
			}
			// customWidth marks the width as user-set — Excel ignores a bare width without it.
			attrs += ` width="${String(width)}" customWidth="1"`
		}
		const hidden = raw.hidden
		if (hidden !== undefined && typeof hidden !== "boolean") {
			sheetInvalid(sheetName, `${what}.hidden must be a boolean`)
		}
		if (hidden === true) attrs += ' hidden="1"'
		if (width !== undefined || hidden === true) entries.push(`<col${attrs}/>`)
	}
	return entries.length > 0 ? `<cols>${entries.join("")}</cols>` : ""
}

// Per-row `ht`/`hidden` attributes, keyed by 1-based row number. Rows whose properties all
// normalize away are dropped; the survivors may belong to rows with no cells at all.
function rowAttrsMap(
	sheetName: string,
	rowProperties: Readonly<Record<number, RowProps>> | undefined,
): Map<number, string> {
	const out = new Map<number, string>()
	if (rowProperties === undefined) return out
	if (!isPlainRecord(rowProperties)) sheetInvalid(sheetName, "rowProperties must be an object")
	for (const key of Object.keys(rowProperties)) {
		const rowNum = Number(key)
		if (!Number.isInteger(rowNum) || rowNum < 1 || rowNum > MAX_ROW) {
			sheetInvalid(
				sheetName,
				`rowProperties key "${key}" is not a row number within Excel's grid`,
			)
		}
		const raw = (rowProperties as Record<string, unknown>)[key]
		const what = `rowProperties[${key}]`
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`)
		checkGeoKeys(sheetName, what, raw, ["height", "hidden"])
		let attrs = ""
		const height = raw.height
		if (height !== undefined) {
			if (
				typeof height !== "number" ||
				!Number.isFinite(height) ||
				height <= 0 ||
				height > MAX_ROW_HEIGHT
			) {
				sheetInvalid(sheetName, `${what}.height must be a number in (0, ${MAX_ROW_HEIGHT}]`)
			}
			// customHeight marks the height as user-set, mirroring customWidth on columns.
			attrs += ` ht="${String(height)}" customHeight="1"`
		}
		const hidden = raw.hidden
		if (hidden !== undefined && typeof hidden !== "boolean") {
			sheetInvalid(sheetName, `${what}.hidden must be a boolean`)
		}
		if (hidden === true) attrs += ' hidden="1"'
		if (attrs !== "") out.set(rowNum, attrs)
	}
	return out
}

// <sheetViews> with a frozen <pane>. For state="frozen", xSplit/ySplit are whole column/row
// counts; topLeftCell is the first scrollable cell and activePane the quadrant the cursor lives
// in — Excel expects all three to be consistent.
function sheetViewsXml(sheetName: string, freeze: FreezePane | undefined): string {
	if (freeze === undefined) return ""
	if (!isPlainRecord(freeze)) sheetInvalid(sheetName, "freeze must be an object")
	checkGeoKeys(sheetName, "freeze", freeze, ["rows", "cols"])
	const validate = (value: unknown, what: string, limit: number): number => {
		if (value === undefined) return 0
		if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value >= limit) {
			sheetInvalid(sheetName, `freeze.${what} must be an integer in [0, ${limit})`)
		}
		return value
	}
	const rows = validate(freeze.rows, "rows", MAX_ROW)
	const cols = validate(freeze.cols, "cols", MAX_COL)
	if (rows === 0 && cols === 0) return "" // freezing nothing is no freeze
	const splits = (cols > 0 ? ` xSplit="${cols}"` : "") + (rows > 0 ? ` ySplit="${rows}"` : "")
	const topLeft = formatRef({ col: cols + 1, row: rows + 1 })
	const activePane = rows > 0 && cols > 0 ? "bottomRight" : rows > 0 ? "bottomLeft" : "topRight"
	return (
		'<sheetViews><sheetView workbookViewId="0">' +
		`<pane${splits} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>` +
		"</sheetView></sheetViews>"
	)
}

export interface WorksheetResult {
	readonly xml: string
}

/** Build the worksheet XML for one sheet, interning cell styles into the shared registry. */
export function worksheetXml(
	sheet: SheetInput,
	date1904: boolean,
	styles: StyleRegistry,
): WorksheetResult {
	const rows = sheet.rows
	// Geometry validates up front (and contributes rows below): a bad column/row/freeze spec must
	// surface before any cell work.
	const cols = colsXml(sheet.name, sheet.columns)
	const rowAttrs = rowAttrsMap(sheet.name, sheet.rowProperties)
	const sheetViews = sheetViewsXml(sheet.name, sheet.freeze)

	// 0 means "unset" — no populated cell seen yet (columns/rows are 1-based, so 0 is a safe sentinel).
	let minRow = 0
	let maxRow = 0
	let minCol = 0
	let maxCol = 0
	// (rowNum, xml) pairs: cell rows arrive in ascending order; property-only rows are merged in
	// afterwards, then the whole set is sorted so <sheetData> stays ascending.
	const rowXmls: [number, string][] = []

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
		if (cellXmls.length > 0) {
			const attrs = rowAttrs.get(rowNum) ?? ""
			rowAttrs.delete(rowNum) // consumed — whatever remains becomes a property-only row
			rowXmls.push([rowNum, `<row r="${rowNum}"${attrs}>${cellXmls.join("")}</row>`])
		}
	}

	// Rows that carry height/hidden but no cells still exist in the file, as cell-less <row>
	// elements. They do not extend the dimension (Excel's dimension covers content, not geometry).
	for (const [rowNum, attrs] of rowAttrs) {
		rowXmls.push([rowNum, `<row r="${rowNum}"${attrs}/>`])
	}
	rowXmls.sort((a, b) => a[0] - b[0])

	// Bounding box of the populated cells, in A1 notation. An entirely empty sheet is "A1" (Excel's
	// convention); a single cell collapses to that one ref rather than a degenerate "X:X" range.
	const dimension =
		minRow === 0
			? "A1"
			: minRow === maxRow && minCol === maxCol
				? formatRef({ col: minCol, row: minRow })
				: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`

	// Schema order within <worksheet>: dimension, sheetViews, cols, sheetData. The geometry blocks
	// are empty strings when unused, so a geometry-free sheet emits the exact pre-F4.5 bytes.
	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"><dimension ref="${dimension}"/>${sheetViews}${cols}<sheetData>${rowXmls
		.map(([, x]) => x)
		.join("")}</sheetData></worksheet>`
	return { xml }
}
