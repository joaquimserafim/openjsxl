import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { openZip } from '../../zip'
import { isDateFormatCode, parseStyles } from '../styles'

describe('isDateFormatCode', () => {
	it('detects date/time format codes', () => {
		for (const code of [
			'mm-dd-yy',
			'yyyy-mm-dd',
			'd/m/yyyy',
			'h:mm:ss',
			'[h]:mm:ss', // elapsed time with tokens outside the brackets
			'[h]', // elapsed hours, wholly bracketed
			'[mm]:[ss]', // elapsed minutes:seconds, wholly bracketed
			'[ss].00',
			'm/d/yy h:mm',
			'[$-409]mmmm d, yyyy', // locale prefix in brackets
			'dddd',
		]) {
			expect(isDateFormatCode(code), code).toBe(true)
		}
	})

	it('rejects non-date format codes', () => {
		for (const code of [
			'General',
			'0.00',
			'0%',
			'#,##0.00',
			'[Red]-0.0;[Blue]0.0', // colour brackets only
			'"day"0', // the date letters are inside a quoted literal
			'\\d0', // escaped d is a literal, not a token
			'@', // text
		]) {
			expect(isDateFormatCode(code), code).toBe(false)
		}
	})
})

describe('parseStyles', () => {
	const xml = `<styleSheet>
	<numFmts count="2">
		<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
		<numFmt numFmtId="165" formatCode="0.000"/>
	</numFmts>
	<cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>
	<cellXfs count="5">
		<xf numFmtId="0"/>
		<xf numFmtId="14" applyNumberFormat="1"/>
		<xf numFmtId="2"/>
		<xf numFmtId="164" applyNumberFormat="1"/>
		<xf numFmtId="165" applyNumberFormat="1"/>
	</cellXfs>
</styleSheet>`
	const styles = parseStyles(xml)

	it('treats built-in date number formats as dates', () => {
		expect(styles.isDateStyle(1)).toBe(true) // numFmtId 14 (mm-dd-yy)
	})

	it('treats custom date number formats as dates', () => {
		expect(styles.isDateStyle(3)).toBe(true) // numFmtId 164 -> "yyyy-mm-dd"
	})

	it('treats numeric formats as not dates', () => {
		expect(styles.isDateStyle(0)).toBe(false) // General
		expect(styles.isDateStyle(2)).toBe(false) // numFmtId 2 (0.00)
		expect(styles.isDateStyle(4)).toBe(false) // numFmtId 165 -> "0.000"
	})

	it('defaults an omitted style index to 0', () => {
		expect(styles.isDateStyle(undefined)).toBe(false)
	})

	it('reads xf only from cellXfs, not cellStyleXfs', () => {
		// cellStyleXfs has one xf (numFmtId 0); if it leaked into the index, the cellXfs
		// entries would shift and isDateStyle(1) would no longer be the date format.
		expect(styles.isDateStyle(1)).toBe(true)
	})

	it('returns false for an out-of-range style index', () => {
		expect(styles.isDateStyle(99)).toBe(false)
	})
})

describe('parseStyles — real basic.xlsx', () => {
	it('flags the date-styled cell format', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		const styles = parseStyles(new TextDecoder().decode(await zip.read('xl/styles.xml')))
		expect(styles.isDateStyle(0)).toBe(false)
		expect(styles.isDateStyle(1)).toBe(true) // C1 uses s="1" -> numFmtId 14
	})
})
