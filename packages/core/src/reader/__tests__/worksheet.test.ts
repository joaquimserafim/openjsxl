import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import type { StyleTable } from '../../ooxml'
import type { Cell } from '../../types'
import { openZip } from '../../zip'
import { type Row, readRows } from '../worksheet'

const ctx = { sharedStrings: ['hello', 'world'] }

const sheet = (body: string) =>
	`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`

const rows = (body: string, sharedStrings: string[] = ctx.sharedStrings): Row[] => [
	...readRows(sheet(body), { sharedStrings }),
]

/** Flatten every yielded cell into a ref -> cell map for direct assertions. */
const byRef = (rs: Row[]): Map<string, Cell> => {
	const map = new Map<string, Cell>()
	for (const row of rs) for (const cell of row.cells) map.set(cell.ref, cell)
	return map
}

describe('readRows — cell typing', () => {
	it('decodes each t variant in a row', () => {
		const cells = byRef(
			rows(
				'<row r="1">' +
					'<c r="A1" t="s"><v>0</v></c>' +
					'<c r="B1"><v>42</v></c>' +
					'<c r="C1" t="b"><v>1</v></c>' +
					'<c r="D1" t="e"><v>#REF!</v></c>' +
					'<c r="E1" t="inlineStr"><is><t>inline</t></is></c>' +
					'<c r="F1" t="str"><f>A1</f><v>cached</v></c>' +
					'<c r="G1"><f>B1*2</f><v>84</v></c>' +
					'</row>',
			),
		)
		expect(cells.get('A1')).toEqual({ ref: 'A1', type: 'string', value: 'hello' })
		expect(cells.get('B1')).toEqual({ ref: 'B1', type: 'number', value: 42 })
		expect(cells.get('C1')).toEqual({ ref: 'C1', type: 'boolean', value: true })
		expect(cells.get('D1')).toEqual({ ref: 'D1', type: 'error', value: '#REF!' })
		expect(cells.get('E1')).toEqual({ ref: 'E1', type: 'string', value: 'inline' })
		expect(cells.get('F1')).toEqual({ ref: 'F1', type: 'string', value: 'cached' })
		expect(cells.get('G1')).toEqual({ ref: 'G1', type: 'number', value: 84 })
	})

	it('concatenates inline rich-text runs and excludes phonetic guides', () => {
		const cells = byRef(
			rows(
				'<row r="1">' +
					'<c r="A1" t="inlineStr"><is><r><t>Hello </t></r><r><t>World</t></r></is></c>' +
					'<c r="B1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh></is></c>' +
					'</row>',
			),
		)
		expect(cells.get('A1')?.value).toBe('Hello World')
		expect(cells.get('B1')?.value).toBe('漢字')
	})

	it('preserves significant whitespace inside inline text', () => {
		const cells = byRef(
			rows('<row r="1"><c r="A1" t="inlineStr"><is><t>  a b </t></is></c></row>'),
		)
		expect(cells.get('A1')?.value).toBe('  a b ')
	})

	it('reads an empty or value-less cell as empty', () => {
		const cells = byRef(rows('<row r="1"><c r="A1"/><c r="B1"></c><c r="C1" t="s"/></row>'))
		expect(cells.get('A1')).toEqual({ ref: 'A1', type: 'empty', value: null })
		expect(cells.get('B1')?.type).toBe('empty')
		expect(cells.get('C1')?.type).toBe('empty')
	})

	it('reads an explicitly empty string/inline value as "" (present, not blank)', () => {
		const cells = byRef(
			rows(
				'<row r="1">' +
					'<c r="A1" t="str"><f>""</f><v></v></c>' +
					'<c r="B1" t="inlineStr"><is><t></t></is></c>' +
					'</row>',
			),
		)
		expect(cells.get('A1')).toEqual({ ref: 'A1', type: 'string', value: '' })
		expect(cells.get('B1')).toEqual({ ref: 'B1', type: 'string', value: '' })
	})

	it('reads only the channel its type selects, ignoring a stray sibling element', () => {
		const cells = byRef(
			rows(
				'<row r="1">' +
					// <v> is the boolean channel; the stray <is> must not pollute it.
					'<c r="A1" t="b"><is><t>x</t></is><v>1</v></c>' +
					// <v> is the error channel; the stray <is> must not append to it.
					'<c r="B1" t="e"><is><t>junk</t></is><v>#N/A</v></c>' +
					// <is> is the inline channel; the stray <v> must not append to it.
					'<c r="C1" t="inlineStr"><is><t>real</t></is><v>99</v></c>' +
					'</row>',
			),
		)
		expect(cells.get('A1')).toEqual({ ref: 'A1', type: 'boolean', value: true })
		expect(cells.get('B1')).toEqual({ ref: 'B1', type: 'error', value: '#N/A' })
		expect(cells.get('C1')).toEqual({ ref: 'C1', type: 'string', value: 'real' })
	})
})

describe('readRows — layout', () => {
	it('keeps sparse cells without filling gaps', () => {
		const [row] = rows('<row r="1"><c r="A1"><v>1</v></c><c r="D1"><v>4</v></c></row>')
		expect(row?.cells.map((c) => c.ref)).toEqual(['A1', 'D1'])
	})

	it('yields out-of-order cells in document order, keyed by ref', () => {
		const [row] = rows('<row r="1"><c r="C1"><v>3</v></c><c r="A1"><v>1</v></c></row>')
		expect(row?.cells.map((c) => c.ref)).toEqual(['C1', 'A1'])
		expect(byRef([row as Row]).get('A1')?.value).toBe(1)
	})

	it('assigns columns positionally when r is absent', () => {
		const [row] = rows(
			'<row r="5"><c><v>1</v></c><c><v>2</v></c><c r="E5"><v>5</v></c><c><v>6</v></c></row>',
		)
		expect(row?.cells.map((c) => c.ref)).toEqual(['A5', 'B5', 'E5', 'F5'])
	})

	it('numbers rows positionally when r is absent', () => {
		const rs = rows('<row><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>')
		expect(rs.map((r) => r.index)).toEqual([1, 2])
	})

	it('emits multiple rows in order, including an explicit empty row', () => {
		const rs = rows(
			'<row r="1"><c r="A1"><v>1</v></c></row><row r="2"/><row r="3"><c r="A3"><v>3</v></c></row>',
		)
		expect(rs.map((r) => r.index)).toEqual([1, 2, 3])
		expect(rs[1]?.cells).toEqual([])
	})

	it('ignores markup outside <sheetData>', () => {
		const xml =
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
		const rs = [...readRows(xml, ctx)]
		expect(rs).toHaveLength(1)
		expect(byRef(rs).get('A1')?.value).toBe(1)
	})
})

describe('readRows — real basic.xlsx', () => {
	it('reads every cell of sheet1 with the right type', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		const xml = new TextDecoder().decode(await zip.read('xl/worksheets/sheet1.xml'))
		const cells = byRef([...readRows(xml, ctx)])

		// No style table is passed here, so date-styled C1 still reads as a number — date
		// promotion is exercised below and in the openXlsx integration test.
		expect(cells.get('A1')).toEqual({ ref: 'A1', type: 'string', value: 'hello' })
		expect(cells.get('B1')).toEqual({ ref: 'B1', type: 'number', value: 42 })
		expect(cells.get('C1')).toEqual({ ref: 'C1', type: 'number', value: 43831 })
		expect(cells.get('D1')).toEqual({ ref: 'D1', type: 'boolean', value: true })
		expect(cells.get('E1')).toEqual({ ref: 'E1', type: 'number', value: 84 })
		expect(cells.get('A2')).toEqual({ ref: 'A2', type: 'string', value: 'world' })
		expect(cells.get('B2')).toEqual({ ref: 'B2', type: 'number', value: 3.14159 })
	})

	it('promotes a date-styled number to a date when a style table is supplied', () => {
		const styles: StyleTable = { isDateStyle: (i) => i === 1 }
		const cells = byRef([
			...readRows(
				sheet('<row r="1"><c r="A1" s="1"><v>43831</v></c><c r="B1"><v>42</v></c></row>'),
				{ sharedStrings: [], styles },
			),
		])
		expect(cells.get('A1')).toEqual({
			ref: 'A1',
			type: 'date',
			value: new Date(Date.UTC(2020, 0, 1)),
		})
		expect(cells.get('B1')).toEqual({ ref: 'B1', type: 'number', value: 42 })
	})
})
