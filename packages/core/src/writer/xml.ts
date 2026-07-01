// XML text serialization for the writer — the encode side of the reader's entity decoding.
//
// OOXML content is UTF-8, so the only characters that MUST be escaped are the ones that would
// otherwise be read as markup: `&`, `<`, `>` in text, plus `"` inside a double-quoted attribute.
// `&` is replaced first so the `&` we introduce for the others isn't re-escaped. We do not emit
// numeric character references for anything else — the bytes go out as UTF-8.

export function escapeText(s: string): string {
	return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function escapeAttr(s: string): string {
	return escapeText(s).replace(/"/g, '&quot;')
}

// Whether `s` can be serialized into XML text/attributes without corrupting the document or the
// data. Rejected: the C0 control characters except tab/LF/CR and the non-characters U+FFFE/U+FFFF
// (none are legal in an XML 1.0 document, even as a numeric reference — Excel would refuse to open
// the file), and lone surrogates (TextEncoder silently swaps them for U+FFFD, losing data). A
// char-code scan rather than a regex, so the source carries no literal control characters and pairs
// surrogates in a single pass. Well-formed astral characters (emoji) and tab/LF/CR are fine.
//
// (Excel round-trips control characters via the _xHHHH_ escape; supporting that is a possible future
// enhancement — it would need matching decode support in the reader.)
export function isXmlSafe(s: string): boolean {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i)
		if (c < 0x20) {
			if (c !== 0x09 && c !== 0x0a && c !== 0x0d) return false
		} else if (c === 0xfffe || c === 0xffff) {
			return false
		} else if (c >= 0xd800 && c <= 0xdbff) {
			// High surrogate — must be immediately followed by a low surrogate to form a valid pair.
			const next = s.charCodeAt(i + 1)
			if (!(next >= 0xdc00 && next <= 0xdfff)) return false
			i++ // consume the paired low surrogate
		} else if (c >= 0xdc00 && c <= 0xdfff) {
			// Low surrogate not preceded by a high one — a lone surrogate.
			return false
		}
	}
	return true
}

// A string whose meaningful leading/trailing whitespace would be stripped by an XML reader needs
// `xml:space="preserve"` on its element, or Excel drops those spaces on load. We flag it whenever
// trimming would change the string (covers leading/trailing spaces, tabs, and newlines).
export function needsPreserve(s: string): boolean {
	return s !== s.trim()
}

/** The ` xml:space="preserve"` attribute (with leading space) when `s` needs it, else `''`. */
export function preserveAttr(s: string): string {
	return needsPreserve(s) ? ' xml:space="preserve"' : ''
}
