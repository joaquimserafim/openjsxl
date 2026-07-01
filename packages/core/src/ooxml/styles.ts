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
	/**
	 * The number-format code applied at this `s` index — a custom code (`<numFmts>`) or a
	 * built-in one (e.g. `"0.00%"`, `"mm-dd-yy"`). `undefined` when the id is a locale/reserved
	 * built-in with no portable code, or the index is out of range.
	 */
	formatCode(styleIndex: number | undefined): string | undefined
}

// Built-in number formats (ECMA-376 §18.8.30) with a fixed, non-locale code. The locale and
// reserved ids (23–36, 41–44, 50–58) have no portable code and resolve to undefined; date
// detection still recognises those via isBuiltinDateId, we just don't fabricate their string.
const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
	0: 'General',
	1: '0',
	2: '0.00',
	3: '#,##0',
	4: '#,##0.00',
	5: '"$"#,##0_);("$"#,##0)',
	6: '"$"#,##0_);[Red]("$"#,##0)',
	7: '"$"#,##0.00_);("$"#,##0.00)',
	8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
	9: '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'mm-dd-yy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yy h:mm',
	37: '#,##0_);(#,##0)',
	38: '#,##0_);[Red](#,##0)',
	39: '#,##0.00_);(#,##0.00)',
	40: '#,##0.00_);[Red](#,##0.00)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	47: 'mmss.0',
	48: '##0.0E+0',
	49: '@',
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

	function formatCode(styleIndex: number | undefined): string | undefined {
		// An omitted `s` means style 0, the implicit default format. A custom code for the id
		// wins over the built-in table (a file may redefine one); unknown ids stay undefined.
		const numFmtId = cellFormatIds[styleIndex ?? 0]
		if (numFmtId === undefined) return undefined
		return customFormats.get(numFmtId) ?? BUILTIN_FORMATS[numFmtId]
	}

	return { isDateStyle, formatCode }
}
