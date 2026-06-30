import { describe, expect, it } from 'vitest'
import { parseWorkbook } from '../workbook'

describe('parseWorkbook', () => {
	it('lists sheets in order with name, r:id, and visibility', () => {
		const xml = `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<sheets>
		<sheet name="First" sheetId="1" r:id="rId1"/>
		<sheet name="Second" sheetId="2" r:id="rId2"/>
	</sheets>
</workbook>`
		expect(parseWorkbook(xml)).toEqual([
			{ name: 'First', rid: 'rId1', visible: true },
			{ name: 'Second', rid: 'rId2', visible: true },
		])
	})

	it('marks hidden and very-hidden sheets not visible', () => {
		const xml =
			'<workbook><sheets>' +
			'<sheet name="A" r:id="rId1" state="visible"/>' +
			'<sheet name="B" r:id="rId2" state="hidden"/>' +
			'<sheet name="C" r:id="rId3" state="veryHidden"/>' +
			'</sheets></workbook>'
		expect(parseWorkbook(xml).map((s) => s.visible)).toEqual([true, false, false])
	})

	it('finds the relationship id under any namespace prefix', () => {
		const xml = '<workbook><sheets><sheet name="A" x:id="rId7"/></sheets></workbook>'
		expect(parseWorkbook(xml)[0]?.rid).toBe('rId7')
	})

	it('skips sheets missing a name or relationship id', () => {
		const xml =
			'<workbook><sheets><sheet name="A"/><sheet r:id="rId2"/><sheet name="C" r:id="rId3"/></sheets></workbook>'
		expect(parseWorkbook(xml).map((s) => s.name)).toEqual(['C'])
	})
})
