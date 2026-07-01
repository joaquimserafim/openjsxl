import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { openXlsx } from '../workbook'
import { parseDimension } from '../worksheet'

// Worksheet dimension (F2.3). The `<dimension ref>` element declares the sheet's used range.
// It is optional, so the accessor is `string | undefined`. Verified against real output:
// basic.xlsx (a range), any_sheets.xlsx (a single-cell dimension, and a sheet with none).

describe('Worksheet.dimension — real fixtures', () => {
	it('reads a range dimension', async () => {
		const wb = await openXlsx(await loadFixture('basic.xlsx'))
		expect(wb.sheet('Sheet1').dimension).toBe('A1:E2')
	})

	it('reads a single-cell dimension', async () => {
		const wb = await openXlsx(await loadFixture('any_sheets.xlsx'))
		expect(wb.sheet('Hidden').dimension).toBe('A1')
	})

	it('is undefined for a sheet that declares no dimension', async () => {
		const wb = await openXlsx(await loadFixture('any_sheets.xlsx'))
		const chart = wb.sheet('Chart')
		expect(chart.dimension).toBeUndefined()
		expect(chart.dimension).toBeUndefined() // cached: still undefined on the second read
	})
})

describe('parseDimension — units', () => {
	it('returns the ref of the dimension element', () => {
		expect(parseDimension('<worksheet><dimension ref="A1:C3"/><sheetData/></worksheet>')).toBe(
			'A1:C3',
		)
	})

	it('tolerates a namespace prefix', () => {
		expect(parseDimension('<x:dimension ref="B2"/>')).toBe('B2')
	})

	it('is undefined when there is no dimension element', () => {
		expect(
			parseDimension('<worksheet><sheetData><row r="1"/></sheetData></worksheet>'),
		).toBeUndefined()
	})

	it('is undefined for a missing or empty ref', () => {
		expect(parseDimension('<dimension/>')).toBeUndefined()
		expect(parseDimension('<dimension ref=""/>')).toBeUndefined()
	})
})
