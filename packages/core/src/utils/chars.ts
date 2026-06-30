// Small, pure character predicates shared across the parsing layers. Kept inside core
// for now; may graduate to a standalone @openjsxl/utils package later.

/** True for the four characters XML treats as whitespace: space, tab, CR, LF. */
export function isWhitespace(ch: string | undefined): boolean {
	return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r'
}
