import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { openXlsx } from '../workbook'

// F4.1 end-to-end — Worksheet.style(ref) against a REAL producer. openpyxl-styled.xlsx was
// authored by openpyxl 3.1.5 (see fixtures/data/README.md), so these assertions are non-circular:
// they check our reader against another implementation's output, not our own writer's.

async function styledSheet() {
	return (await openXlsx(await loadFixture('openpyxl-styled.xlsx'))).sheet('Styled')
}

describe('Worksheet.style — openpyxl-authored fixture', () => {
	it('reads fonts: bold+color+name+size, italic+underline, strike, indexed color', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('A1')?.font).toEqual({
			name: 'Arial',
			size: 14,
			bold: true,
			color: { rgb: 'FFFF0000' },
		})
		expect(sheet.style('B1')?.font).toEqual({ italic: true, underline: 'single' })
		expect(sheet.style('A5')?.font).toEqual({ strike: true })
		expect(sheet.style('B5')?.font).toEqual({ color: { indexed: 10 } })
	})

	it('keeps a theme+tint font color raw', async () => {
		const sheet = await styledSheet()
		const color = sheet.style('B2')?.font?.color
		expect(color).toEqual({ theme: 4, tint: 0.3999755851924192 })
	})

	it('reads solid and patterned fills', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('A2')?.fill).toEqual({
			patternType: 'solid',
			fgColor: { rgb: 'FFFFFF00' },
		})
		expect(sheet.style('C2')?.fill).toEqual({
			patternType: 'lightGray',
			fgColor: { rgb: 'FF00B050' },
		})
	})

	it('reads per-edge borders with optional colors', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('A3')?.border).toEqual({
			left: { style: 'dashed', color: { rgb: 'FF0070C0' } },
			top: { style: 'thin', color: { rgb: 'FF000000' } },
			bottom: { style: 'double' },
		})
	})

	it('reads alignment, including rotation', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('B3')?.alignment).toEqual({
			horizontal: 'center',
			vertical: 'top',
			wrapText: true,
			indent: 2,
		})
		expect(sheet.style('C3')?.alignment).toEqual({ textRotation: 45 })
	})

	it('returns number formats as code strings — built-in and custom', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('A4')).toEqual({ numberFormat: '#,##0.00' }) // built-in id 4
		expect(sheet.style('B4')).toEqual({ numberFormat: '0.00%' }) // built-in id 10
		expect(sheet.style('C4')?.numberFormat).toBe('"kg" 0.0') // custom, id >= 164
	})

	it('reads a fully-loaded style with every component at once', async () => {
		const sheet = await styledSheet()
		const style = sheet.style('C5')
		expect(style?.numberFormat).toBe('0.0')
		expect(style?.font).toEqual({ bold: true, italic: true, size: 9 })
		expect(style?.fill).toEqual({ patternType: 'solid', fgColor: { rgb: 'FFDDEBF7' } })
		expect(style?.border).toEqual({ top: { style: 'medium' }, bottom: { style: 'medium' } })
		expect(style?.alignment).toEqual({ horizontal: 'right' })
	})

	it('returns undefined for unstyled and absent cells', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('C1')).toBeUndefined() // written plain, xf 0
		expect(sheet.style('Z99')).toBeUndefined() // absent cell
	})

	it('is reference-stable per distinct format', async () => {
		const sheet = await styledSheet()
		expect(sheet.style('A1')).toBe(sheet.style('A1'))
	})
})

describe('Worksheet.style — agreement with numberFormat on inherited styles', () => {
	// col-row-styles.xlsx (#22): column B carries a date column-default (numFmtId 14) and row 3 a
	// percent row-default. style() resolves through the SAME effective-style map as
	// numberFormat(), so the two must tell one story for inherited styles too.
	it('resolves column and row default styles identically on both accessors', async () => {
		const sheet = (await openXlsx(await loadFixture('col-row-styles.xlsx'))).sheet('Sheet1')

		// B1: bare cell in the date-default column.
		expect(sheet.style('B1')?.numberFormat).toBe('mm-dd-yy')
		expect(sheet.numberFormat('B1')).toBe('mm-dd-yy')

		// A3: bare cell in the percent-default row.
		expect(sheet.style('A3')?.numberFormat).toBe('0.00%')
		expect(sheet.numberFormat('A3')).toBe('0.00%')

		// A1: unstyled — numberFormat resolves to the default 'General'; style() deliberately
		// answers "no style" instead (the default format is not a per-cell style).
		expect(sheet.numberFormat('A1')).toBe('General')
		expect(sheet.style('A1')).toBeUndefined()
	})
})
