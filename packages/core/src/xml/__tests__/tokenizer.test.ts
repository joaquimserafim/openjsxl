import { describe, expect, it } from 'vitest'
import { tokenize, type XmlToken } from '../tokenizer'

function tokens(xml: string): XmlToken[] {
	return [...tokenize(xml)]
}

function textRuns(xml: string): string[] {
	return tokens(xml)
		.filter((t) => t.kind === 'text')
		.map((t) => t.value)
}

describe('tokenize', () => {
	it('emits open/text/close for a simple element', () => {
		expect(tokens('<a>hello</a>')).toEqual([
			{ kind: 'open', name: 'a', attrs: {}, selfClosing: false },
			{ kind: 'text', value: 'hello' },
			{ kind: 'close', name: 'a' },
		])
	})

	it('parses double- and single-quoted attributes', () => {
		expect(tokens('<c r="A1" s=\'1\'>x</c>')).toEqual([
			{ kind: 'open', name: 'c', attrs: { r: 'A1', s: '1' }, selfClosing: false },
			{ kind: 'text', value: 'x' },
			{ kind: 'close', name: 'c' },
		])
	})

	it('marks self-closing tags and emits no close', () => {
		expect(tokens('<c r="A1"/>')).toEqual([
			{ kind: 'open', name: 'c', attrs: { r: 'A1' }, selfClosing: true },
		])
		expect(tokens('<border/>')).toEqual([
			{ kind: 'open', name: 'border', attrs: {}, selfClosing: true },
		])
	})

	it('keeps namespaced prefixes as part of the literal name', () => {
		expect(tokens('<sheet r:id="rId1" xml:space="preserve"/>')).toEqual([
			{
				kind: 'open',
				name: 'sheet',
				attrs: { 'r:id': 'rId1', 'xml:space': 'preserve' },
				selfClosing: true,
			},
		])
	})

	it('decodes entities in text and attribute values', () => {
		expect(tokens('<a t="x &amp; y">a &lt; b</a>')).toEqual([
			{ kind: 'open', name: 'a', attrs: { t: 'x & y' }, selfClosing: false },
			{ kind: 'text', value: 'a < b' },
			{ kind: 'close', name: 'a' },
		])
	})

	it('keeps ">" and "/" inside quoted attribute values', () => {
		// A literal ">" in a quoted value must not end the tag (e.g. cached formulas).
		expect(tokens('<a b="x>y"/>')).toEqual([
			{ kind: 'open', name: 'a', attrs: { b: 'x>y' }, selfClosing: true },
		])
		// A "/" in a quoted value must not trigger self-close early (e.g. rels targets).
		expect(tokens('<rel Target="worksheets/sheet1.xml"/>')).toEqual([
			{
				kind: 'open',
				name: 'rel',
				attrs: { Target: 'worksheets/sheet1.xml' },
				selfClosing: true,
			},
		])
	})

	it('recovers a stray self-close slash without dropping a following attribute', () => {
		expect(tokens('<a x="1"/ y="2">')).toEqual([
			{ kind: 'open', name: 'a', attrs: { x: '1', y: '2' }, selfClosing: true },
		])
	})

	it('preserves significant whitespace in text', () => {
		expect(tokens('<t xml:space="preserve"> </t>')).toEqual([
			{ kind: 'open', name: 't', attrs: { 'xml:space': 'preserve' }, selfClosing: false },
			{ kind: 'text', value: ' ' },
			{ kind: 'close', name: 't' },
		])
	})

	it('emits inter-element whitespace as text (pretty-printed input)', () => {
		expect(textRuns('<row>\n\t<c/>\n</row>')).toEqual(['\n\t', '\n'])
	})

	it('skips the xml prolog and comments', () => {
		expect(tokens('<?xml version="1.0" encoding="UTF-8"?><a/>')).toEqual([
			{ kind: 'open', name: 'a', attrs: {}, selfClosing: true },
		])
		expect(tokens('<a><!-- note --><b/></a>')).toEqual([
			{ kind: 'open', name: 'a', attrs: {}, selfClosing: false },
			{ kind: 'open', name: 'b', attrs: {}, selfClosing: true },
			{ kind: 'close', name: 'a' },
		])
	})

	it('strips a leading UTF-8 BOM', () => {
		expect(tokens('﻿<?xml version="1.0"?><sst/>')).toEqual([
			{ kind: 'open', name: 'sst', attrs: {}, selfClosing: true },
		])
	})

	it('treats CDATA as literal text and emits one run per character-data span', () => {
		expect(tokens('<a><![CDATA[<raw> & stuff]]></a>')).toEqual([
			{ kind: 'open', name: 'a', attrs: {}, selfClosing: false },
			{ kind: 'text', value: '<raw> & stuff' },
			{ kind: 'close', name: 'a' },
		])
		// Character data is split across a CDATA boundary; consumers must concatenate.
		expect(textRuns('<t>a<![CDATA[b]]>c</t>')).toEqual(['a', 'b', 'c'])
	})

	it('recovers an unescaped "<" in text instead of inventing an element', () => {
		const result = tokens('<a>3 < 5 and x</a>')
		expect(result.some((t) => t.kind === 'open' && t.name === '')).toBe(false)
		expect(result.at(-1)).toEqual({ kind: 'close', name: 'a' })
		expect(textRuns('<a>3 < 5 and x</a>').join('')).toBe('3 < 5 and x')
	})

	it('tokenizes a realistic worksheet row', () => {
		const xml =
			'<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c>' +
			'<c r="C1" s="1"><v>43831</v></c><c r="D1" t="b"><v>1</v></c>' +
			'<c r="E1"><f>B1*2</f><v>84</v></c></row>'
		const result = tokens(xml)
		expect(result[0]).toEqual({
			kind: 'open',
			name: 'row',
			attrs: { r: '1' },
			selfClosing: false,
		})
		expect(result[1]).toEqual({
			kind: 'open',
			name: 'c',
			attrs: { r: 'A1', t: 's' },
			selfClosing: false,
		})
		expect(result[2]).toEqual({ kind: 'open', name: 'v', attrs: {}, selfClosing: false })
		expect(result[3]).toEqual({ kind: 'text', value: '0' })
		expect(result[4]).toEqual({ kind: 'close', name: 'v' })
		// the cached formula cell carries both <f> and <v>
		const text = textRuns(xml)
		expect(text).toContain('B1*2')
		expect(text).toContain('84')
	})

	it('tokenizes a shared strings table', () => {
		const xml =
			'<sst count="2" uniqueCount="2"><si><t>hello</t></si><si><t>world</t></si></sst>'
		expect(textRuns(xml)).toEqual(['hello', 'world'])
	})
})
