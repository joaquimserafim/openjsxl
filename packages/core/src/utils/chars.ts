// Small, pure character predicates shared across the parsing layers. Kept inside core
// for now; may graduate to a standalone @openjsxl/utils package later.

/** True for the four characters XML treats as whitespace: space, tab, CR, LF. */
export function isWhitespace(ch: string | undefined): boolean {
	return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

// Whether `s` can live in XML text/attributes without corrupting the document or the data.
// Rejected: the C0 control characters except tab/LF/CR and the non-characters U+FFFE/U+FFFF
// (none are legal in an XML 1.0 document, even as a numeric reference — Excel refuses to open a
// file carrying one), and lone surrogates (TextEncoder silently swaps them for U+FFFD, losing
// data). A char-code scan rather than a regex, so the source carries no literal control
// characters and surrogates pair in a single pass. Tab/LF/CR and well-formed astral characters
// (emoji) are fine. Shared by the writer (reject unsafe input) and the reader's style model
// (degrade unsafe producer strings, so the bridge only ever carries writable values).
export function isXmlSafe(s: string): boolean {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c < 0x20) {
			if (c !== 0x09 && c !== 0x0a && c !== 0x0d) return false;
		} else if (c === 0xfffe || c === 0xffff) {
			return false;
		} else if (c >= 0xd800 && c <= 0xdbff) {
			// High surrogate — must be immediately followed by a low surrogate to form a pair.
			const next = s.charCodeAt(i + 1);
			if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
			i++; // consume the paired low surrogate
		} else if (c >= 0xdc00 && c <= 0xdfff) {
			// Low surrogate not preceded by a high one — a lone surrogate.
			return false;
		}
	}
	return true;
}
