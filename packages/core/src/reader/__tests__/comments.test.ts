import { existsSync } from 'node:fs'
import { loadFixture, loadLocalFixture, localFixturePath } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { openXlsx } from '../workbook'
import { parseComments } from '../worksheet'

// Comments (F2.3). They live in a separate part (xl/commentsN.xml) linked from the worksheet
// rels: an <authors> list plus a <commentList> whose <comment ref authorId> holds rich text.
// parseComments concatenates the <t> runs and resolves the author. The parser is covered by
// units here; real Excel output is verified against a local-only Apache-2.0 POI file (skipped
// when absent — see packages/fixtures/local).

describe('parseComments — units', () => {
	it('resolves the author and concatenates rich-text runs', () => {
		const xml =
			'<comments><authors><author>Ada</author></authors>' +
			'<commentList><comment ref="B2" authorId="0"><text>' +
			'<r><rPr><b/></rPr><t>bold </t></r><r><t xml:space="preserve">and plain</t></r>' +
			'</text></comment></commentList></comments>'
		expect(parseComments(xml)).toEqual([{ ref: 'B2', author: 'Ada', text: 'bold and plain' }])
	})

	it('reads plain <t> comments and a second author', () => {
		const xml =
			'<comments><authors><author>A</author><author>B</author></authors><commentList>' +
			'<comment ref="A1" authorId="0"><text><t>hi</t></text></comment>' +
			'<comment ref="C3" authorId="1"><text><t>yo</t></text></comment>' +
			'</commentList></comments>'
		expect(parseComments(xml)).toEqual([
			{ ref: 'A1', author: 'A', text: 'hi' },
			{ ref: 'C3', author: 'B', text: 'yo' },
		])
	})

	it('omits the author when the authorId resolves to nothing', () => {
		const xml =
			'<comments><authors><author>A</author></authors>' +
			'<commentList><comment ref="A1" authorId="9"><text><t>x</t></text></comment></commentList></comments>'
		expect(parseComments(xml)).toEqual([{ ref: 'A1', text: 'x' }])
	})

	it('omits the author when a comment has no authorId (no default to author 0)', () => {
		const xml =
			'<comments><authors><author>A</author></authors>' +
			'<commentList><comment ref="A1"><text><t>x</t></text></comment></commentList></comments>'
		expect(parseComments(xml)).toEqual([{ ref: 'A1', text: 'x' }])
	})

	it('tolerates a namespace prefix and an empty comment', () => {
		const xml =
			'<x:comments><x:authors><x:author>A</x:author></x:authors>' +
			'<x:commentList><x:comment ref="A1" authorId="0"><x:text/></x:comment></x:commentList></x:comments>'
		expect(parseComments(xml)).toEqual([{ ref: 'A1', author: 'A', text: '' }])
	})

	it('is empty when there are no comments', () => {
		expect(parseComments('<comments><authors/><commentList/></comments>')).toEqual([])
	})
})

describe('Worksheet.comments — no comments part', () => {
	it('is empty for a sheet without comments', async () => {
		const wb = await openXlsx(await loadFixture('basic.xlsx'))
		expect(wb.sheet('Sheet1').comments).toEqual([])
	})
})

// Real Excel comment output — Apache-2.0 POI fixtures kept local-only (git/npm-ignored).
// Runs when present; skipped on a fresh clone or in CI.
const poi = 'SimpleWithComments.xlsx'

describe.skipIf(!existsSync(localFixturePath(poi)))(
	'Worksheet.comments — real Excel (local)',
	() => {
		it('reads authored, multi-line comments from genuine Excel output', async () => {
			const wb = await openXlsx(await loadLocalFixture(poi))
			expect(wb.sheet('Sheet1').comments).toEqual([
				{ ref: 'B1', author: 'Yegor Kozlov', text: 'Yegor Kozlov:\r\nfirst cell' },
				{ ref: 'B2', author: 'Yegor Kozlov', text: 'Yegor Kozlov:\r\nsecond cell' },
				{ ref: 'B3', author: 'Yegor Kozlov', text: 'Yegor Kozlov:\r\nthird cell' },
			])
		})
	},
)
