import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import type { DecodeContext } from '../../ooxml'
import { openXlsx, streamSheetRows } from '../workbook'
import { parseCellStyles, readRows } from '../worksheet'

// #22 — a cell's format can be set at the column (`<col style>`) or row (`<row s customFormat>`)
// level, not only on the cell. The reader must resolve the effective style with the precedence
//   cell `s` > row default (customFormat) > column default > style 0
// so both date detection and numberFormat() see the format a producer actually applied.

describe('parseCellStyles — effective style precedence', () => {
	// Inline worksheet XML exercises the resolver directly, with no fixture file: the map holds
	// each addressed cell's EFFECTIVE cellXfs index (absent ⇒ resolves to style 0).
	const styleOf = (xml: string, ref: string) => parseCellStyles(xml).get(ref)
	const sheet = (inner: string) => `<worksheet>${inner}</worksheet>`

	it('inherits a column default when the cell omits its own s', () => {
		const xml = sheet(
			'<cols><col min="1" max="1" style="7"/></cols><sheetData><row r="1"><c r="A1"/></row></sheetData>',
		)
		expect(styleOf(xml, 'A1')).toBe(7)
	})

	it('applies a column default only within its [min,max] range', () => {
		const xml = sheet(
			'<cols><col min="2" max="3" style="7"/></cols><sheetData><row r="1"><c r="A1"/><c r="B1"/><c r="D1"/></row></sheetData>',
		)
		expect(styleOf(xml, 'A1')).toBeUndefined() // col 1, outside 2..3
		expect(styleOf(xml, 'B1')).toBe(7) // col 2, inside
		expect(styleOf(xml, 'D1')).toBeUndefined() // col 4, outside
	})

	it('inherits a row default only when customFormat is set', () => {
		const gated = sheet(
			'<sheetData><row r="1" s="5" customFormat="1"><c r="A1"/></row></sheetData>',
		)
		expect(styleOf(gated, 'A1')).toBe(5)

		// A row `s` WITHOUT customFormat is bookkeeping, not a format — it must NOT restyle cells.
		const ungated = sheet('<sheetData><row r="1" s="5"><c r="A1"/></row></sheetData>')
		expect(styleOf(ungated, 'A1')).toBeUndefined()
	})

	it('lets the cell s win over both row and column defaults', () => {
		const xml = sheet(
			'<cols><col min="1" max="1" style="7"/></cols><sheetData><row r="1" s="5" customFormat="1"><c r="A1" s="3"/></row></sheetData>',
		)
		expect(styleOf(xml, 'A1')).toBe(3)
	})

	it('lets the row default win over the column default', () => {
		const xml = sheet(
			'<cols><col min="1" max="1" style="7"/></cols><sheetData><row r="1" s="5" customFormat="1"><c r="A1"/></row></sheetData>',
		)
		expect(styleOf(xml, 'A1')).toBe(5)
	})

	it('does not carry a row default into a following non-customFormat row', () => {
		const xml = sheet(
			'<sheetData><row r="1" s="5" customFormat="1"><c r="A1"/></row><row r="2"><c r="A2"/></row></sheetData>',
		)
		expect(styleOf(xml, 'A1')).toBe(5)
		expect(styleOf(xml, 'A2')).toBeUndefined()
	})

	it('keys a cell with no r by its positional ref, matching the assembler', () => {
		// A cell without `r` is positioned one past the previous cell — parseCellStyles must key
		// it the same way the assembler does, or numberFormat() and cell() would disagree.
		const xml = sheet(
			'<cols><col min="2" max="2" style="7"/></cols><sheetData><row r="1"><c r="A1"/><c/></row></sheetData>',
		)
		expect(styleOf(xml, 'B1')).toBe(7) // the no-r cell, positioned at column 2, inherits it
	})
})

describe('inherited styles — no-r cells agree across accessors', () => {
	// Regression (adversarial review): a cell that omits its `r` is positioned by the assembler
	// (so cell()/date-detection sees it) but was skipped by parseCellStyles (so numberFormat()
	// missed it) — the two public accessors disagreed for the same synthesized ref.
	const ctx: DecodeContext = {
		sharedStrings: [],
		styles: {
			isDateStyle: (i) => i === 7,
			formatCode: (i) => (i === 7 ? 'mm-dd-yy' : 'General'),
		},
	}

	it('resolves the same ref to the same (date) style on both paths', () => {
		const xml =
			'<worksheet><cols><col min="2" max="2" style="7"/></cols><sheetData>' +
			'<row r="1"><c r="A1"><v>1</v></c><c><v>45000</v></c></row></sheetData></worksheet>'

		// Assembler side (drives cell()/rows()): the no-r cell lands at B1 and reads as a date.
		const b1 = [...readRows(xml, ctx)].flatMap((row) => row.cells).find((c) => c.ref === 'B1')
		expect(b1?.type).toBe('date')

		// numberFormat side: parseCellStyles keys the same B1 with the inherited date style.
		expect(parseCellStyles(xml).get('B1')).toBe(7)
	})
})

describe('inherited styles — end to end (col-row-styles.xlsx)', () => {
	// Column B is date-formatted; row 3 is a customFormat percent row. See the fixture spec in
	// packages/fixtures/scripts/generate.mjs.
	it('detects dates from a column default (cells omit their own s)', async () => {
		const wb = await openXlsx(await loadFixture('col-row-styles.xlsx'))
		const s = wb.sheet('Sheet1')
		expect(s.cell('B1').type).toBe('date') // inherited from column B
		expect(s.cell('B2').type).toBe('date') // inherited from column B
		expect(s.cell('C1').type).toBe('date') // explicit per-cell date (control)
		expect(s.cell('A1').type).toBe('number') // unstyled → plain number
	})

	it('resolves number formats through column and row inheritance', async () => {
		const wb = await openXlsx(await loadFixture('col-row-styles.xlsx'))
		const s = wb.sheet('Sheet1')
		expect(s.numberFormat('B1')).toBe('mm-dd-yy') // column default
		expect(s.numberFormat('B2')).toBe('mm-dd-yy')
		expect(s.numberFormat('C1')).toBe('mm-dd-yy') // explicit
		expect(s.numberFormat('A3')).toBe('0.00%') // row-3 customFormat default
		expect(s.numberFormat('A1')).toBe('General') // style 0
	})

	it('lets a cell style win over its row default', async () => {
		const wb = await openXlsx(await loadFixture('col-row-styles.xlsx'))
		const s = wb.sheet('Sheet1')
		// B3 sits in the percent row 3 but sets its own date s — the cell wins.
		expect(s.cell('B3').type).toBe('date')
		expect(s.numberFormat('B3')).toBe('mm-dd-yy')
		// A3 has no own s → inherits the row's percent, which is not a date.
		expect(s.cell('A3').type).toBe('number')
	})

	it('applies column inheritance on the streaming path too', async () => {
		const byRef = new Map<string, string>()
		for await (const row of streamSheetRows(await loadFixture('col-row-styles.xlsx'))) {
			for (const cell of row.cells) byRef.set(cell.ref, cell.type)
		}
		expect(byRef.get('B1')).toBe('date') // column default resolved while streaming
		expect(byRef.get('B2')).toBe('date')
		expect(byRef.get('B3')).toBe('date') // cell s wins
		expect(byRef.get('A3')).toBe('number') // row percent is not a date
	})
})
