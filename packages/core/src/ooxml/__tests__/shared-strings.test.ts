import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { openZip } from '../../zip'
import { parseSharedStrings } from '../shared-strings'

describe('parseSharedStrings', () => {
	it('reads plain <si><t> items in order', () => {
		const xml = `<sst count="2" uniqueCount="2">
	<si><t>hello</t></si>
	<si><t>world</t></si>
</sst>`
		expect(parseSharedStrings(xml)).toEqual(['hello', 'world'])
	})

	it('concatenates rich-text runs into a single string', () => {
		const xml = `<sst>
	<si>
		<r><rPr><b/></rPr><t>Hello </t></r>
		<r><t>World</t></r>
	</si>
</sst>`
		expect(parseSharedStrings(xml)).toEqual(['Hello World'])
	})

	it('preserves significant whitespace inside <t>', () => {
		const xml =
			'<sst><si><t xml:space="preserve">  leading</t></si><si><r><t>a </t></r><r><t> b</t></r></si></sst>'
		expect(parseSharedStrings(xml)).toEqual(['  leading', 'a  b'])
	})

	it('ignores layout whitespace between elements', () => {
		const xml = `<sst>
	<si>
		<t>x</t>
	</si>
</sst>`
		expect(parseSharedStrings(xml)).toEqual(['x'])
	})

	it('decodes entities in text', () => {
		const xml = '<sst><si><t>a &amp; b &lt;c&gt;</t></si></sst>'
		expect(parseSharedStrings(xml)).toEqual(['a & b <c>'])
	})

	it('treats empty and self-closing items as empty strings', () => {
		const xml = '<sst><si><t></t></si><si><t/></si><si/></sst>'
		expect(parseSharedStrings(xml)).toEqual(['', '', ''])
	})

	it('excludes phonetic (<rPh>) reading text from the value', () => {
		const xml =
			'<sst><si><r><t>漢字</t></r><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></si></sst>'
		expect(parseSharedStrings(xml)).toEqual(['漢字'])
	})

	it('tolerates namespace prefixes on the elements', () => {
		const xml = '<x:sst><x:si><x:t>ns</x:t></x:si></x:sst>'
		expect(parseSharedStrings(xml)).toEqual(['ns'])
	})

	it('excludes a non-empty <phoneticPr> reading and keeps the normal empty form', () => {
		const leaky = '<sst><si><t>main</t><phoneticPr><t>BAD</t></phoneticPr></si></sst>'
		expect(parseSharedStrings(leaky)).toEqual(['main'])
		const normal = '<sst><si><t>x</t><phoneticPr fontId="1"/></si></sst>'
		expect(parseSharedStrings(normal)).toEqual(['x'])
	})

	it('returns an empty table for an empty <sst>', () => {
		expect(parseSharedStrings('<sst count="0" uniqueCount="0"/>')).toEqual([])
	})

	// Misnested markup must not drop or shift the index of well-formed neighbours: we keep
	// one entry per <si> start so cell lookups by index stay aligned.
	describe('misnested <si> recovery', () => {
		it('finalizes an open item when a new <si> opens inside it', () => {
			const xml = '<sst><si><t>a</t><si><t>b</t></si><si><t>c</t></si></sst>'
			expect(parseSharedStrings(xml)).toEqual(['a', 'b', 'c'])
		})

		it('finalizes an open item when a self-closing <si/> appears inside it', () => {
			const xml = '<sst><si><t>a</t><si/><t>b</t></si><si><t>c</t></si></sst>'
			expect(parseSharedStrings(xml)).toEqual(['a', '', 'c'])
		})

		it('keeps one entry per <si> start for stacked self-closing items', () => {
			expect(parseSharedStrings('<sst><si><si/><si/><t>z</t></si></sst>')).toEqual([
				'',
				'',
				'',
			])
		})
	})
})

describe('shared strings — real basic.xlsx', () => {
	it('matches the fixture sst exactly', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		const xml = new TextDecoder().decode(await zip.read('xl/sharedStrings.xml'))
		expect(parseSharedStrings(xml)).toEqual(['hello', 'world'])
	})
})
