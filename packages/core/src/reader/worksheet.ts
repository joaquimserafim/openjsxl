import { type DecodeContext, decodeCell, formatRef, parseRef } from '../ooxml'
import type { Cell } from '../types'
import { localName } from '../utils'
import { createXmlStream, tokenize, type XmlToken } from '../xml'

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
	index: number
	/** Cells present in the row, in document order. Gaps are simply absent (sparse). */
	cells: Cell[]
}

function safeColumn(ref: string): number | undefined {
	try {
		return parseRef(ref).col
	} catch {
		return undefined
	}
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

	let inCell = false
	let cellRef = ''
	let cellType: string | undefined
	let cellStyle: number | undefined // the `s` attribute (index into cellXfs)
	let cellIsInline = false // type === 'inlineStr': value lives in <is>, not <v>
	let cellValue = ''
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

		if (token.kind === 'open') {
			const name = localName(token.name)

			if (name === 'sheetData') {
				if (!token.selfClosing) inSheetData = true
				return out
			}
			if (!inSheetData) return out

			if (name === 'row') {
				flushCell()
				if (inRow) out.push({ index: rowIndex, cells })
				const r = token.attrs.r
				const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN
				rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1
				lastRow = rowIndex
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

			if (name === 'c') {
				flushCell()
				const r = token.attrs.r
				if (r !== undefined) {
					cellRef = r
					const col = safeColumn(r)
					if (col !== undefined) lastCol = col
				} else {
					lastCol += 1
					cellRef = formatRef({ col: lastCol, row: rowIndex })
				}
				cellType = token.attrs.t
				const s = token.attrs.s
				cellStyle = s === undefined ? undefined : Number(s)
				cellIsInline = cellType === 'inlineStr'
				cellValue = ''
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
				if (name === 'is') {
					hasValue = true
					if (!token.selfClosing) inInline = true
				} else if (name === 't') {
					if (inInline && !token.selfClosing) textDepth++
				} else if (name === 'rPh' || name === 'phoneticPr') {
					if (inInline && !token.selfClosing) phoneticDepth++
				}
			} else if (name === 'v') {
				hasValue = true
				if (!token.selfClosing) inValue = true
			}
			return out
		}

		if (token.kind === 'text') {
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
		if (name === 'sheetData') {
			flushCell()
			if (inRow) {
				out.push({ index: rowIndex, cells })
				inRow = false
			}
			inSheetData = false
			return out
		}
		if (name === 'row') {
			if (inRow) {
				flushCell()
				out.push({ index: rowIndex, cells })
				inRow = false
			}
			return out
		}
		if (!inCell) return out
		if (name === 'c') flushCell()
		else if (name === 'v') inValue = false
		else if (name === 'is') inInline = false
		else if (name === 't') {
			if (textDepth > 0) textDepth--
		} else if (name === 'rPh' || name === 'phoneticPr') {
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

/** Read rows from a fully in-memory worksheet string. */
export function* readRows(xml: string, ctx: DecodeContext): Generator<Row> {
	const assembler = createRowAssembler(ctx)
	for (const token of tokenize(xml)) yield* assembler.push(token)
	yield* assembler.flush()
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
		if (text === '') continue
		for (const token of xml.push(text)) yield* assembler.push(token)
	}
	const tail = decoder.decode() // finalize any pending multi-byte sequence
	if (tail !== '') for (const token of xml.push(tail)) yield* assembler.push(token)
	for (const token of xml.flush()) yield* assembler.push(token)
	yield* assembler.flush()
}
