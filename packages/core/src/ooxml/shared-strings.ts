import { localName } from '../utils'
import { tokenize } from '../xml'

// The shared string table (xl/sharedStrings.xml) deduplicates text across the workbook:
// every `<si>` is one string, and a cell of type `s` stores a zero-based index into the
// table rather than repeating the text. Parsing it once up front turns those indices into
// O(1) lookups while reading sheets.
//
// An `<si>` is either a single `<t>` (plain text) or a sequence of `<r>` runs each with its
// own `<t>` (rich text with per-run formatting we don't keep) — the value is every run's
// text concatenated in order. Two subtleties:
//   * Whitespace is significant. The tokenizer emits text verbatim, and we only collect
//     text that sits *inside* a `<t>`, so layout whitespace between elements never leaks in
//     and `xml:space="preserve"` needs no special handling.
//   * Phonetic guides (`<rPh>` and `<phoneticPr>`) carry an alternate reading, not part of
//     the displayed value, so their `<t>` text is excluded — matching openpyxl/Excel.
//
// The tokenizer does not validate structure, so misnested markup is possible. We hold one
// invariant for robustness: exactly one table entry per `<si>` start. A new `<si>`/`<si/>`
// seen while an item is still open finalizes that open item first, so a stray tag can never
// silently drop or shift the index of a well-formed neighbour (indices feed cell lookups,
// where an off-by-one would mislabel unrelated cells).

export function parseSharedStrings(xml: string): string[] {
	const strings: string[] = []

	let inItem = false // within an <si>
	let current = '' // accumulated text for the current <si>
	let textDepth = 0 // open <t> elements (text counts only when > 0)
	let phoneticDepth = 0 // open <rPh>/<phoneticPr> elements (their text is excluded)

	for (const token of tokenize(xml)) {
		if (token.kind === 'open') {
			const name = localName(token.name)
			if (name === 'si') {
				// Finalize an item left open by misnested markup before starting the next.
				if (inItem) {
					strings.push(current)
					inItem = false
				}
				if (token.selfClosing) {
					strings.push('') // an empty self-closed item
				} else {
					inItem = true
					current = ''
					textDepth = 0
					phoneticDepth = 0
				}
			} else if (inItem && !token.selfClosing) {
				if (name === 't') textDepth++
				else if (name === 'rPh' || name === 'phoneticPr') phoneticDepth++
			}
		} else if (token.kind === 'text') {
			if (inItem && textDepth > 0 && phoneticDepth === 0) current += token.value
		} else {
			// close
			const name = localName(token.name)
			if (!inItem) continue
			if (name === 't') {
				if (textDepth > 0) textDepth--
			} else if (name === 'rPh' || name === 'phoneticPr') {
				if (phoneticDepth > 0) phoneticDepth--
			} else if (name === 'si') {
				strings.push(current)
				inItem = false
			}
		}
	}

	return strings
}
