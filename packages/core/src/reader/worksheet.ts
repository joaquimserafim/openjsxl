import { type DecodeContext, decodeCell, formatRef, parseRef, type Relationship } from "../ooxml"
import type { Cell, Comment, Hyperlink } from "../types"
import { localName, relationshipId } from "../utils"
import { createXmlStream, tokenize, type XmlToken } from "../xml"

// Turn a worksheet part (xl/worksheets/sheetN.xml) into rows of typed cells. We walk the
// tokenizer event stream rather than building a DOM, so peak memory tracks one row, not the
// whole sheet. The same row state machine drives both the in-memory `readRows` (over a full
// string) and the constant-memory `streamRows` (over decompressed chunks) — see F2.2.
//
// Cells and rows are sparse and may appear out of order, so each cell carries its own A1
// ref and callers key by that ref — never by position in the array. When a `<c>` omits its
// `r` attribute (some streaming writers do), the column is taken positionally: one past the
// previous cell, exactly as the spec prescribes.
//
// Robustness mirrors the lower layers: the tokenizer never throws and does not validate
// structure, so this code must not corrupt a well-formed row because of a malformed
// neighbour. A new `<row>`/`<c>` opening while one is still unclosed finalizes the open one
// first, and a row left open at end-of-input (truncated file) is still emitted.

export interface Row {
	/** 1-based row index — from `<row r>`, or positional when the attribute is absent. */
	readonly index: number
	/** Cells present in the row, in document order. Gaps are simply absent (sparse). */
	readonly cells: readonly Cell[]
}

function safeColumn(ref: string): number | undefined {
	try {
		return parseRef(ref).col
	} catch {
		return undefined
	}
}

// A cell's format can be set at three levels, not just on the cell. `<c s>` is per-cell, but a
// whole column can be formatted with `<col … style>` and a whole row with `<row s customFormat>`.
// The effective style index resolves in this precedence (SpreadsheetML §18.3.1.13/.73/.4):
//
//     cell's own `s`  >  row default (customFormat)  >  column default (<col>)  >  style 0
//
// Without this, a date column whose cells omit `s` reads as plain numbers, and numberFormat()
// returns the wrong code — the common case where a producer formats a column once instead of
// per cell. All three levels are indices into the SAME cellXfs table, so the resolved index
// drops straight into StyleTable.isDateStyle/formatCode with no further translation.

// A column-default style: `<col … style>` applies to every cell in columns [min, max] that does
// not set its own `s`. (`style` is the column's default cell format — an index into cellXfs.)
interface ColumnStyle {
	readonly min: number
	readonly max: number
	readonly style: number
}

function columnStyleFromToken(attrs: Readonly<Record<string, string>>): ColumnStyle | undefined {
	// A <col> with no `style` sets width/visibility only — it contributes no default format.
	if (attrs.style === undefined) return undefined
	const style = Number(attrs.style)
	const min = Number(attrs.min)
	const max = Number(attrs.max)
	if (!Number.isInteger(style) || !Number.isInteger(min) || !Number.isInteger(max))
		return undefined
	return { min, max, style }
}

// A row's default style is honored ONLY when `customFormat="1"`. The spec is explicit: `<row>`'s
// `s` is "only applied if the customFormat attribute is '1'". Writers routinely emit a bookkeeping
// `s` on ordinary rows, so without this gate we would restyle cells the file never meant to format.
function rowDefaultStyleFromToken(attrs: Readonly<Record<string, string>>): number | undefined {
	if (attrs.customFormat !== "1" && attrs.customFormat !== "true") return undefined
	if (attrs.s === undefined) return undefined
	const s = Number(attrs.s)
	return Number.isInteger(s) ? s : undefined
}

// Resolve the precedence above for one cell. A present-but-unparseable cell `s` still counts as
// "the cell declared its own style" — it opts out of inherited defaults and falls back to style 0,
// rather than silently borrowing the row/column format.
function effectiveStyle(
	ownS: string | undefined,
	rowDefault: number | undefined,
	col: number | undefined,
	columns: readonly ColumnStyle[],
): number | undefined {
	if (ownS !== undefined) {
		const s = Number(ownS)
		return Number.isInteger(s) ? s : undefined
	}
	if (rowDefault !== undefined) return rowDefault
	if (col !== undefined) {
		// First covering range wins. Ranges don't overlap in a valid file; first-match keeps
		// resolution deterministic even if a malformed file declares a column twice.
		const range = columns.find((c) => col >= c.min && col <= c.max)
		if (range !== undefined) return range.style
	}
	return undefined
}

interface RowAssembler {
	/** Advance the state machine by one token; returns any rows that completed. */
	push(token: XmlToken): Row[]
	/** Emit a row left open at end of input (truncated file). */
	flush(): Row[]
}

function createRowAssembler(ctx: DecodeContext): RowAssembler {
	let inSheetData = false
	let inRow = false
	let lastRow = 0
	let rowIndex = 0
	let cells: Cell[] = []
	let lastCol = 0

	// Column defaults from `<cols>` (which precedes `<sheetData>`, so it is fully seen before any
	// cell) and the current row's default (customFormat) — both feed effectiveStyle for cells
	// that omit their own `s`. See the ColumnStyle/effectiveStyle notes above.
	const columns: ColumnStyle[] = []
	let rowDefault: number | undefined

	let inCell = false
	let cellRef = ""
	let cellType: string | undefined
	let cellStyle: number | undefined // the `s` attribute (index into cellXfs)
	let cellIsInline = false // type === 'inlineStr': value lives in <is>, not <v>
	let cellValue = ""
	let hasValue = false // the value channel's element was present (even if empty)
	let inValue = false // inside <v>
	let inInline = false // inside <is>
	let textDepth = 0 // open <t> within <is>
	let phoneticDepth = 0 // open <rPh>/<phoneticPr> within <is> (excluded from the value)

	// Finalize the open cell into the current row. A no-op when no cell is open.
	const flushCell = () => {
		if (!inCell) return
		cells.push(
			decodeCell(
				{
					ref: cellRef,
					type: cellType,
					value: hasValue ? cellValue : undefined,
					style: cellStyle,
				},
				ctx,
			),
		)
		inCell = false
	}

	function push(token: XmlToken): Row[] {
		const out: Row[] = []

		if (token.kind === "open") {
			const name = localName(token.name)

			if (name === "sheetData") {
				if (!token.selfClosing) inSheetData = true
				return out
			}
			// `<col>` lives in `<cols>`, a sibling that precedes `<sheetData>` — capture its
			// default style before the inSheetData guard below would skip it.
			if (name === "col") {
				const cs = columnStyleFromToken(token.attrs)
				if (cs !== undefined) columns.push(cs)
				return out
			}
			if (!inSheetData) return out

			if (name === "row") {
				flushCell()
				if (inRow) out.push({ index: rowIndex, cells })
				const r = token.attrs.r
				const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN
				rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1
				lastRow = rowIndex
				// A non-customFormat row overwrites this back to undefined — no stale carryover.
				rowDefault = rowDefaultStyleFromToken(token.attrs)
				cells = []
				lastCol = 0
				if (token.selfClosing) {
					out.push({ index: rowIndex, cells })
					inRow = false
				} else {
					inRow = true
				}
				return out
			}
			if (!inRow) return out

			if (name === "c") {
				flushCell()
				const r = token.attrs.r
				// The cell's column index, needed to resolve a `<col>` default style. Known even
				// when `r` is absent (positional: one past the previous cell).
				let col: number | undefined
				if (r !== undefined) {
					cellRef = r
					col = safeColumn(r)
					if (col !== undefined) lastCol = col
				} else {
					lastCol += 1
					col = lastCol
					cellRef = formatRef({ col: lastCol, row: rowIndex })
				}
				cellType = token.attrs.t
				// Effective style: cell `s` → row default → column default → style 0.
				cellStyle = effectiveStyle(token.attrs.s, rowDefault, col, columns)
				cellIsInline = cellType === "inlineStr"
				cellValue = ""
				hasValue = false
				inValue = false
				inInline = false
				textDepth = 0
				phoneticDepth = 0
				if (token.selfClosing) {
					cells.push(
						decodeCell(
							{ ref: cellRef, type: cellType, value: undefined, style: cellStyle },
							ctx,
						),
					)
				} else {
					inCell = true
				}
				return out
			}
			if (!inCell) return out

			// A cell's value lives in exactly one channel, picked by its type: inline strings
			// in <is>/<t>, everything else in <v>. Gate on the type so a stray element from
			// the other channel can't pollute the value. Mark the value present as soon as its
			// element opens, so an explicit but empty <v></v> or <is><t></t></is> reads as ""
			// rather than collapsing to a blank cell.
			if (cellIsInline) {
				if (name === "is") {
					hasValue = true
					if (!token.selfClosing) inInline = true
				} else if (name === "t") {
					if (inInline && !token.selfClosing) textDepth++
				} else if (name === "rPh" || name === "phoneticPr") {
					if (inInline && !token.selfClosing) phoneticDepth++
				}
			} else if (name === "v") {
				hasValue = true
				if (!token.selfClosing) inValue = true
			}
			return out
		}

		if (token.kind === "text") {
			const collect = cellIsInline
				? inInline && textDepth > 0 && phoneticDepth === 0
				: inValue
			if (inCell && collect) {
				cellValue += token.value
				hasValue = true
			}
			return out
		}

		// close
		const name = localName(token.name)
		if (name === "sheetData") {
			flushCell()
			if (inRow) {
				out.push({ index: rowIndex, cells })
				inRow = false
			}
			inSheetData = false
			return out
		}
		if (name === "row") {
			if (inRow) {
				flushCell()
				out.push({ index: rowIndex, cells })
				inRow = false
			}
			return out
		}
		if (!inCell) return out
		if (name === "c") flushCell()
		else if (name === "v") inValue = false
		else if (name === "is") inInline = false
		else if (name === "t") {
			if (textDepth > 0) textDepth--
		} else if (name === "rPh" || name === "phoneticPr") {
			if (phoneticDepth > 0) phoneticDepth--
		}
		return out
	}

	function flush(): Row[] {
		flushCell()
		if (inRow) {
			inRow = false
			return [{ index: rowIndex, cells }]
		}
		return []
	}

	return { push, flush }
}

/**
 * The comments a worksheet carries. Comments live in a separate part (xl/commentsN.xml) that
 * pairs an `<authors>` list with a `<commentList>`; each `<comment ref authorId>` holds rich
 * text in `<text>`. We concatenate the `<t>` runs (dropping formatting, matching how shared
 * strings read) and resolve the authorId against the authors list. A comment with no `ref`, or
 * an authorId that resolves to nothing, still yields text — the author is just omitted.
 */
export function parseComments(xml: string): Comment[] {
	const authors: string[] = []
	const comments: Comment[] = []
	let inAuthors = false
	let authorText: string | undefined // non-undefined while inside an <author>
	let current: { ref: string; authorId: number } | undefined // inside a <comment>
	let inText = false // inside the current comment's <text>
	let tDepth = 0 // open <t> within <text>
	let text = ""

	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name)
			if (name === "authors") {
				if (!token.selfClosing) inAuthors = true
			} else if (name === "author" && inAuthors) {
				if (token.selfClosing) authors.push("")
				else authorText = ""
			} else if (name === "comment") {
				const ref = token.attrs.ref
				// A missing/empty/non-integer authorId is unresolved (-1) so the author is omitted,
				// rather than falling through Number('')===0 to author 0 — don't fabricate an
				// attribution the file never made.
				const rawId = token.attrs.authorId
				const authorId =
					rawId !== undefined && rawId !== "" && Number.isInteger(Number(rawId))
						? Number(rawId)
						: -1
				current = ref !== undefined && ref !== "" ? { ref, authorId } : undefined
				text = ""
				inText = false
				tDepth = 0
			} else if (name === "text" && current !== undefined) {
				if (!token.selfClosing) inText = true
			} else if (name === "t" && inText) {
				if (!token.selfClosing) tDepth++
			}
		} else if (token.kind === "text") {
			if (authorText !== undefined) authorText += token.value
			else if (inText && tDepth > 0) text += token.value
		} else {
			const name = localName(token.name)
			if (name === "authors") inAuthors = false
			else if (name === "author" && authorText !== undefined) {
				authors.push(authorText)
				authorText = undefined
			} else if (name === "t") {
				if (tDepth > 0) tDepth--
			} else if (name === "text") inText = false
			else if (name === "comment" && current !== undefined) {
				const author = authors[current.authorId]
				comments.push({
					ref: current.ref,
					...(author !== undefined && author !== "" ? { author } : {}),
					text,
				})
				current = undefined
			}
		}
	}
	return comments
}

/**
 * The worksheet's declared used range — the `<dimension ref>` (e.g. "A1:E10", or a single
 * cell). Optional in OOXML: undefined when the producer omits it, or when it is empty. The
 * element sits near the top of the part, so the scan returns as soon as it is found.
 */
export function parseDimension(xml: string): string | undefined {
	for (const token of tokenize(xml)) {
		if (token.kind === "open" && localName(token.name) === "dimension") {
			const ref = token.attrs.ref
			if (ref !== undefined && ref !== "") return ref
		}
	}
	return undefined
}

/**
 * Map each cell to its EFFECTIVE style index, for resolving per-cell number formats. Effective
 * means the same precedence date detection uses — cell `s`, else the row default (`<row s
 * customFormat>`), else the column default (`<col … style>`), else style 0. A cell that resolves
 * to no style is omitted (the lookup falls back to style 0 anyway).
 *
 * Addressing MUST match the row assembler so numberFormat() and cell()/rows() agree on the same
 * ref: an explicit `r` is used verbatim; a cell without `r` is positioned one past the previous
 * cell (`formatRef`), exactly as the assembler synthesizes its ref — otherwise a no-`r` cell that
 * inherits a column/row style would read as a date via cell() but as "General" via numberFormat().
 */
export function parseCellStyles(xml: string): Map<string, number> {
	const styles = new Map<string, number>()
	const columns: ColumnStyle[] = []
	let rowDefault: number | undefined
	let lastRow = 0
	let rowIndex = 0
	let lastCol = 0
	for (const token of tokenize(xml)) {
		if (token.kind !== "open") continue
		const name = localName(token.name)
		if (name === "col") {
			const cs = columnStyleFromToken(token.attrs)
			if (cs !== undefined) columns.push(cs)
			continue
		}
		if (name === "row") {
			const r = token.attrs.r
			const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN
			rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1
			lastRow = rowIndex
			rowDefault = rowDefaultStyleFromToken(token.attrs)
			lastCol = 0
			continue
		}
		if (name !== "c") continue
		// Resolve the cell's ref and column with the same positional rule as the assembler.
		const r = token.attrs.r
		let ref: string
		let col: number | undefined
		if (r !== undefined) {
			ref = r
			col = safeColumn(r)
			if (col !== undefined) lastCol = col
		} else {
			lastCol += 1
			col = lastCol
			ref = formatRef({ col: lastCol, row: rowIndex })
		}
		const index = effectiveStyle(token.attrs.s, rowDefault, col, columns)
		if (index !== undefined) styles.set(ref, index)
	}
	return styles
}

/** Read rows from a fully in-memory worksheet string. */
export function* readRows(xml: string, ctx: DecodeContext): Generator<Row> {
	const assembler = createRowAssembler(ctx)
	for (const token of tokenize(xml)) yield* assembler.push(token)
	yield* assembler.flush()
}

/**
 * The merged-cell ranges a worksheet declares, in document order, as A1 ranges (e.g. `A1:B2`).
 * They live in a `<mergeCells>` block after `<sheetData>`, so we scan the token stream for the
 * `<mergeCell ref>` children rather than the row state machine. Refs are returned verbatim as
 * the producer wrote them; a missing or empty `ref` is skipped.
 */
export function parseMergedCells(xml: string): string[] {
	const ranges: string[] = []
	for (const token of tokenize(xml)) {
		if (token.kind === "open" && localName(token.name) === "mergeCell") {
			const ref = token.attrs.ref
			if (ref !== undefined && ref !== "") ranges.push(ref)
		}
	}
	return ranges
}

/**
 * The hyperlinks a worksheet declares, in document order. They live in a `<hyperlinks>` block
 * (a sibling of `<sheetData>`); each `<hyperlink>` carries the covered `ref` plus some of:
 * an `r:id` pointing into the worksheet's own relationships (the external target), an
 * in-workbook `location`, a `tooltip`, and a `display` override. The external target is joined
 * from `rels` here so callers get the resolved URL, not a bare relationship id; a link with no
 * resolvable `r:id` simply has no `target`. A missing/empty `ref` is skipped.
 */
export function parseHyperlinks(xml: string, rels?: Map<string, Relationship>): Hyperlink[] {
	const links: Hyperlink[] = []
	for (const token of tokenize(xml)) {
		if (token.kind !== "open" || localName(token.name) !== "hyperlink") continue
		const ref = token.attrs.ref
		if (ref === undefined || ref === "") continue

		// Built immutably (Hyperlink is readonly): include only the attributes that are present.
		const rid = relationshipId(token.attrs)
		const target = rid !== undefined ? rels?.get(rid)?.target : undefined
		const location = token.attrs.location
		const tooltip = token.attrs.tooltip
		const display = token.attrs.display
		links.push({
			ref,
			...(target !== undefined ? { target } : {}),
			...(location !== undefined && location !== "" ? { location } : {}),
			...(tooltip !== undefined ? { tooltip } : {}),
			...(display !== undefined ? { display } : {}),
		})
	}
	return links
}

/**
 * Read rows from a stream of decompressed worksheet chunks without materializing the part —
 * peak memory tracks one row, not the whole sheet (F2.2). Bytes are decoded with a streaming
 * `TextDecoder` (multi-byte sequences may split across chunks) and fed through the chunk-safe
 * tokenizer.
 */
export async function* streamRows(
	chunks: AsyncIterable<Uint8Array>,
	ctx: DecodeContext,
): AsyncGenerator<Row> {
	const assembler = createRowAssembler(ctx)
	const xml = createXmlStream()
	const decoder = new TextDecoder()

	for await (const bytes of chunks) {
		const text = decoder.decode(bytes, { stream: true })
		if (text === "") continue
		for (const token of xml.push(text)) yield* assembler.push(token)
	}
	const tail = decoder.decode() // finalize any pending multi-byte sequence
	if (tail !== "") for (const token of xml.push(tail)) yield* assembler.push(token)
	for (const token of xml.flush()) yield* assembler.push(token)
	yield* assembler.flush()
}
