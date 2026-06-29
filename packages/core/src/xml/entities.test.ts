import { describe, expect, it } from 'vitest'
import { decodeXmlEntities } from './entities'

describe('decodeXmlEntities', () => {
	it('returns the input unchanged when there is no ampersand', () => {
		expect(decodeXmlEntities('plain text')).toBe('plain text')
	})

	it('decodes the five predefined entities', () => {
		expect(decodeXmlEntities('a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')).toBe(
			'a & b < c > d "e" \'f\'',
		)
	})

	it('decodes decimal character references', () => {
		expect(decodeXmlEntities('line&#10;break')).toBe('line\nbreak')
		expect(decodeXmlEntities('&#65;&#66;&#67;')).toBe('ABC')
	})

	it('decodes hexadecimal character references, including astral code points', () => {
		expect(decodeXmlEntities('&#x41;&#x42;')).toBe('AB')
		expect(decodeXmlEntities('&#x1F600;')).toBe('😀')
	})

	it('handles adjacent entities', () => {
		expect(decodeXmlEntities('&lt;&gt;&amp;')).toBe('<>&')
	})

	it('leaves unknown or malformed entities intact', () => {
		expect(decodeXmlEntities('&unknown;')).toBe('&unknown;')
		expect(decodeXmlEntities('&#xZZ;')).toBe('&#xZZ;') // not hex digits
		expect(decodeXmlEntities('5 & 6')).toBe('5 & 6') // bare ampersand
		expect(decodeXmlEntities('a &amp b')).toBe('a &amp b') // missing semicolon
	})
})
