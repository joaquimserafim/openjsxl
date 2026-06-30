import { localName } from '../utils'
import { tokenize } from '../xml'

// xl/styles.xml is what makes a number a date: a cell's `s` attribute indexes `<cellXfs>`,
// each `<xf>` names a `numFmtId`, and the number format that id resolves to decides whether
// the value is a date/time. We keep only what date detection needs — the numFmtId per
// cell-format and the custom format codes — not the full style record.
//
//   numFmtId < 164  built-in, NOT written into <numFmts>. The date/time ones are fixed by
//                   the spec (14–22, 45–47) plus the locale date/time block (27–36, 50–58).
//   numFmtId ≥ 164  custom, defined in <numFmts> with a formatCode we sniff for date tokens.

export interface StyleTable {
	/** True when the cell format at this `s` index applies a date/time number format. */
	isDateStyle(styleIndex: number | undefined): boolean
}

function isBuiltinDateId(id: number): boolean {
	return (
		(id >= 14 && id <= 22) ||
		(id >= 27 && id <= 36) ||
		(id >= 45 && id <= 47) ||
		(id >= 50 && id <= 58)
	)
}

// A format code is a date/time format when, after removing the parts that can never be date
// tokens — quoted literals ("…"), escaped characters (\x), and bracketed sections ([Red],
// [$-409], [>100]) — one of the date/time letters d m y h s remains.
//
// Elapsed-time tokens — [h], [hh], [mm], [ss] — are an exception: they ARE time formats but
// live inside brackets the NON_TOKEN pass would strip, so a code like "[h]" (or "[mm]:[ss]")
// would otherwise reduce to nothing. Detect them up front. (Ordinary elapsed formats like
// "[h]:mm:ss" already pass via the mm/ss outside the brackets.)
const ELAPSED_TIME = /\[(?:h+|m+|s+)\]/i
const NON_TOKEN = /\[[^\]]*\]|"[^"]*"|\\./g
const DATE_TOKEN = /[dmyhs]/i

export function isDateFormatCode(formatCode: string): boolean {
	if (ELAPSED_TIME.test(formatCode)) return true
	return DATE_TOKEN.test(formatCode.replace(NON_TOKEN, ''))
}

export function parseStyles(xml: string): StyleTable {
	const customFormats = new Map<number, string>()
	const cellFormatIds: number[] = []
	let inNumFmts = false
	let inCellXfs = false // NOT cellStyleXfs — a cell's `s` indexes cellXfs only

	for (const token of tokenize(xml)) {
		if (token.kind === 'text') continue
		const name = localName(token.name)
		if (token.kind === 'open') {
			if (name === 'numFmts') {
				if (!token.selfClosing) inNumFmts = true
			} else if (name === 'cellXfs') {
				if (!token.selfClosing) inCellXfs = true
			} else if (name === 'numFmt' && inNumFmts) {
				const id = Number(token.attrs.numFmtId)
				const code = token.attrs.formatCode
				if (Number.isInteger(id) && code !== undefined) customFormats.set(id, code)
			} else if (name === 'xf' && inCellXfs) {
				const id = Number(token.attrs.numFmtId ?? '0')
				cellFormatIds.push(Number.isInteger(id) ? id : 0)
			}
		} else if (token.kind === 'close') {
			if (name === 'numFmts') inNumFmts = false
			else if (name === 'cellXfs') inCellXfs = false
		}
	}

	function isDateStyle(styleIndex: number | undefined): boolean {
		// An omitted `s` means style 0, the implicit default format.
		const numFmtId = cellFormatIds[styleIndex ?? 0]
		if (numFmtId === undefined) return false
		const custom = customFormats.get(numFmtId)
		return custom !== undefined ? isDateFormatCode(custom) : isBuiltinDateId(numFmtId)
	}

	return { isDateStyle }
}
