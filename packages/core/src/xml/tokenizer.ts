import { isWhitespace } from '../utils'
import { decodeXmlEntities } from './entities'

// Streaming, non-validating tokenizer for the OOXML subset of XML. It yields a flat
// event stream rather than building a DOM — the reason pure-JS parsers can keep up with
// native ones is that they never materialise a tree.
//
// Handles: elements, attributes (single or double quoted), self-closing tags, text,
// CDATA, the `<?xml?>` prolog, and comments. Entity references (F1.1) are decoded in
// element text and attribute values. `xml:space="preserve"` needs no special handling
// here — all text is emitted verbatim (after entity decoding) and significance is the
// consumer's call.
//
// Deliberately NOT supported: DTD validation, namespace resolution (prefixes such as
// `r:id` are part of the literal name), and processing instructions beyond the prolog.
//
// Robustness contract: for ANY input it never throws, hangs, or reads out of bounds, and
// it never invents a phantom element for an unescaped `<`. It does NOT guarantee
// completeness — a truncated part yields a plausible-but-shorter stream with no error
// signal, so callers must validate document structure themselves. Character data is
// emitted as one `text` event per run between markup, so consumers must concatenate
// consecutive `text` events (CDATA, comments, and entities all split a run).

export type XmlToken =
	| {
			readonly kind: 'open'
			readonly name: string
			readonly attrs: Readonly<Record<string, string>>
			readonly selfClosing: boolean
	  }
	| { readonly kind: 'text'; readonly value: string }
	| { readonly kind: 'close'; readonly name: string }

export function* tokenize(xml: string): Generator<XmlToken> {
	const len = xml.length
	// Skip a leading UTF-8 BOM if present (common from LibreOffice and older Excel).
	let i = xml.charCodeAt(0) === 0xfeff ? 1 : 0

	while (i < len) {
		const lt = xml.indexOf('<', i)

		if (lt === -1) {
			const text = xml.slice(i)
			if (text.length > 0) yield { kind: 'text', value: decodeXmlEntities(text) }
			break
		}

		if (lt > i) {
			yield { kind: 'text', value: decodeXmlEntities(xml.slice(i, lt)) }
		}

		const next = xml[lt + 1]

		// <?xml ... ?> prolog / processing instruction
		if (next === '?') {
			const end = xml.indexOf('?>', lt + 2)
			i = end === -1 ? len : end + 2
			continue
		}

		// <!-- comment -->, <![CDATA[ ... ]]>, or other declaration
		if (next === '!') {
			if (xml.startsWith('<!--', lt)) {
				const end = xml.indexOf('-->', lt + 4)
				i = end === -1 ? len : end + 3
				continue
			}
			if (xml.startsWith('<![CDATA[', lt)) {
				const end = xml.indexOf(']]>', lt + 9)
				const content = xml.slice(lt + 9, end === -1 ? len : end)
				if (content.length > 0) yield { kind: 'text', value: content } // CDATA is literal
				i = end === -1 ? len : end + 3
				continue
			}
			const end = xml.indexOf('>', lt + 2)
			i = end === -1 ? len : end + 1
			continue
		}

		// </name> close tag
		if (next === '/') {
			const end = xml.indexOf('>', lt + 2)
			const name = xml.slice(lt + 2, end === -1 ? len : end).trim()
			yield { kind: 'close', name }
			i = end === -1 ? len : end + 1
			continue
		}

		// A '<' not followed by a name-start character is a literal '<' in (malformed or
		// unescaped) text — emit it verbatim rather than inventing a phantom element and
		// swallowing the following close tag.
		if (next === undefined || next === '>' || next === '=' || isWhitespace(next)) {
			yield { kind: 'text', value: '<' }
			i = lt + 1
			continue
		}

		// <name ...> open or self-closing tag
		let j = lt + 1
		const nameStart = j
		while (j < len) {
			const ch = xml[j]
			if (isWhitespace(ch) || ch === '>' || ch === '/') break
			j++
		}
		const name = xml.slice(nameStart, j)
		const attrs: Record<string, string> = {}
		let selfClosing = false

		while (j < len) {
			while (j < len && isWhitespace(xml[j])) j++
			if (j >= len) break

			const ch = xml[j]
			if (ch === '>') {
				j++
				break
			}
			if (ch === '/') {
				// Self-close marker: record it and advance by one, letting the loop consume the
				// real '>'. Do NOT jump to the next '>' — that can swallow following markup.
				selfClosing = true
				j++
				continue
			}

			const attrNameStart = j
			while (j < len) {
				const c = xml[j]
				if (c === '=' || c === '>' || c === '/' || isWhitespace(c)) break
				j++
			}
			const attrName = xml.slice(attrNameStart, j)

			while (j < len && isWhitespace(xml[j])) j++
			if (xml[j] === '=') {
				j++
				while (j < len && isWhitespace(xml[j])) j++
				const quote = xml[j]
				if (quote === '"' || quote === "'") {
					j++
					const close = xml.indexOf(quote, j)
					const end = close === -1 ? len : close
					if (attrName.length > 0) attrs[attrName] = decodeXmlEntities(xml.slice(j, end))
					j = close === -1 ? len : close + 1
				} else {
					// Unquoted value (malformed XML, but be lenient): read to whitespace/end-of-tag.
					const valStart = j
					while (j < len && !isWhitespace(xml[j]) && xml[j] !== '>' && xml[j] !== '/') j++
					if (attrName.length > 0)
						attrs[attrName] = decodeXmlEntities(xml.slice(valStart, j))
				}
			} else if (attrName.length > 0) {
				attrs[attrName] = ''
			}
		}

		yield { kind: 'open', name, attrs, selfClosing }
		i = j
	}
}
