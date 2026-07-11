import { FormulaError } from "./errors";

// Tokenizer for the stored (en-US, ECMA-376 §18.17) formula form. A hand-rolled scanner rather than
// one mega-regex, because the cell/name/function/sheet distinctions need lookahead the parser is best
// placed to make: this layer only splits the text into the smallest unambiguous pieces and leaves
// assembling references (sheet prefixes, ranges, 3-D spans) to the parser.
//
// Whitespace is insignificant here and dropped. Excel's space *intersection* operator (`A1:A5 B1:B5`)
// is therefore not modelled in v0.8 — a named parse-level exclusion; such a formula becomes a typed
// parse error rather than a wrong tree. Every other operator, including the `:` range operator and the
// `,` union operator, is preserved.

export type TokenType =
	| "num" // numeric literal (raw kept in value)
	| "str" // "..." string literal (value already unescaped)
	| "err" // one of the eight #... cell-error literals
	| "sheet" // '...' quoted sheet name (value already unescaped)
	| "name" // identifier: function/defined name, unquoted sheet name, TRUE/FALSE, whole-col letter
	| "cell" // a cell address, `$` markers kept (A1, $A$1)
	| "bracket" // a balanced [...] span, kept raw (structured/external reference material)
	| "op" // an operator or punctuation character (value is the operator text)
	| "eof"; // end of input sentinel

export interface Token {
	readonly type: TokenType;
	readonly value: string;
	/** 0-based start offset into the source. */
	readonly start: number;
	/** 0-based offset one past the token's last character. */
	readonly end: number;
}

// The eight ST_CellErrorType literals (ECMA-376 §18.18.11), longest-first so a startsWith scan can
// never stop at a shorter member that happens to be a prefix (none are, but order makes it robust).
const ERROR_LITERALS = [
	"#GETTING_DATA",
	"#DIV/0!",
	"#VALUE!",
	"#NAME?",
	"#NULL!",
	"#NUM!",
	"#REF!",
	"#N/A",
];

function isDigit(code: number): boolean {
	return code >= 48 && code <= 57; // 0-9
}

function isAsciiLetter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122); // A-Z a-z
}

// A name may start with a letter, `_`, `\`, or any non-ASCII character (Unicode-named sheets/names).
function isNameStart(code: number): boolean {
	return isAsciiLetter(code) || code === 95 || code === 92 || code >= 0x80;
}

// After the first character a name may additionally contain digits and `.`.
function isNameChar(code: number): boolean {
	return isNameStart(code) || isDigit(code) || code === 46; // '.'
}

// Try to read a cell address (`$?[A-Za-z]{1,3}$?[0-9]{1,7}`) at `pos`. Returns the end offset, or
// -1 when the run is not a cell. The trailing boundary is what disambiguates a cell from a longer
// token: `LOG10(` is a function (next char `(`), `Sheet1!` a sheet name (next `!`), `A1B` a name
// (a 4th letter) — all rejected here so they fall through to the name path. Out-of-grid shapes like
// `ZZZ1` are still accepted as cells; bounds are the evaluator's concern, not the lexer's.
function matchCell(s: string, pos: number): number {
	const len = s.length;
	let i = pos;
	if (s.charCodeAt(i) === 36) i++; // optional leading '$'
	let letters = 0;
	while (i < len && isAsciiLetter(s.charCodeAt(i)) && letters < 3) {
		i++;
		letters++;
	}
	if (letters === 0) return -1;
	if (i < len && isAsciiLetter(s.charCodeAt(i))) return -1; // a 4th letter → not a cell
	if (i < len && s.charCodeAt(i) === 36) i++; // optional '$' before the row
	let digits = 0;
	while (i < len && isDigit(s.charCodeAt(i)) && digits < 7) {
		i++;
		digits++;
	}
	if (digits === 0) return -1;
	if (i < len && isDigit(s.charCodeAt(i))) return -1; // an 8th digit → not a cell
	const next = i < len ? s.charCodeAt(i) : -1;
	// A cell cannot be immediately followed by more reference/name material: letter/digit/_/./\,
	// `$`, `(` (function), `!` (sheet name), or `[` (structured ref) all mean this run is something
	// bigger and must be re-read as a name.
	if (
		next !== -1 &&
		(isNameChar(next) || next === 36 || next === 40 || next === 33 || next === 91)
	)
		return -1;
	return i;
}

// Read a `$`-prefixed PARTIAL reference at `pos` (matchCell has already declined a full cell here):
// an absolute whole-column endpoint (`$A`, `$AB`) or absolute whole-row endpoint (`$2`) — the pieces
// of `$A:$B` and `$2:$5`. Returns the end offset, or -1 when the `$` is stray. The unmarked forms
// (`A`, `1`) need no special case: a bare column lexes as a name and a bare row as a number, and the
// parser's `:` operator combines either into a range.
function matchDollarRef(s: string, pos: number): number {
	const len = s.length;
	let i = pos + 1; // past the '$'
	if (i < len && isAsciiLetter(s.charCodeAt(i))) {
		let letters = 0;
		while (i < len && isAsciiLetter(s.charCodeAt(i)) && letters < 3) {
			i++;
			letters++;
		}
		// A trailing digit/name char/`$` would have made this a cell or a longer token — decline.
		if (i < len && (isNameChar(s.charCodeAt(i)) || s.charCodeAt(i) === 36)) return -1;
		return i;
	}
	if (i < len && isDigit(s.charCodeAt(i))) {
		let digits = 0;
		while (i < len && isDigit(s.charCodeAt(i)) && digits < 7) {
			i++;
			digits++;
		}
		if (i < len && isDigit(s.charCodeAt(i))) return -1;
		return i;
	}
	return -1;
}

// Read a numeric literal at `pos` (integer, decimal, or scientific). Returns the end offset. A stray
// `e`/`E` with no exponent digits is not consumed as part of the number (it starts a name instead).
function matchNumber(s: string, pos: number): number {
	const len = s.length;
	let i = pos;
	while (i < len && isDigit(s.charCodeAt(i))) i++;
	if (i < len && s.charCodeAt(i) === 46) {
		i++;
		while (i < len && isDigit(s.charCodeAt(i))) i++;
	}
	if (i < len && (s.charCodeAt(i) === 101 || s.charCodeAt(i) === 69)) {
		let j = i + 1;
		if (j < len && (s.charCodeAt(j) === 43 || s.charCodeAt(j) === 45)) j++; // '+' / '-'
		if (j < len && isDigit(s.charCodeAt(j))) {
			j++;
			while (j < len && isDigit(s.charCodeAt(j))) j++;
			i = j;
		}
	}
	return i;
}

// Read a quoted run delimited by `quote` (`"` string or `'` sheet name), where a doubled quote is a
// literal quote. Returns the unescaped value and the end offset (past the closing quote).
function matchQuoted(s: string, pos: number, quote: number): { value: string; end: number } {
	const len = s.length;
	let i = pos + 1;
	let out = "";
	while (i < len) {
		const c = s.charCodeAt(i);
		if (c === quote) {
			if (i + 1 < len && s.charCodeAt(i + 1) === quote) {
				out += String.fromCharCode(quote);
				i += 2;
				continue;
			}
			return { value: out, end: i + 1 };
		}
		out += s[i];
		i++;
	}
	throw new FormulaError("parse-error", "unterminated quoted literal", { position: pos });
}

// Read a balanced `[...]` span at `pos` (structured/external reference material), returning the end
// offset. Nesting is tracked (`Table1[[#Data],[Amt]]`), and a `'` escapes the next character so an
// escaped `[`/`]` inside a column name does not unbalance the count.
function matchBracket(s: string, pos: number): number {
	const len = s.length;
	let i = pos;
	let depth = 0;
	while (i < len) {
		const c = s.charCodeAt(i);
		if (c === 39) {
			i += 2; // "'" escapes the following char
			continue;
		}
		if (c === 91) depth++;
		else if (c === 93) {
			depth--;
			if (depth === 0) return i + 1;
		}
		i++;
	}
	throw new FormulaError("parse-error", "unbalanced '[' in reference", { position: pos });
}

/** Tokenize a stored-form formula. Throws a typed {@link FormulaError} on a malformed literal. */
export function tokenize(source: string): Token[] {
	const tokens: Token[] = [];
	const len = source.length;
	let i = 0;
	while (i < len) {
		const code = source.charCodeAt(i);

		// Whitespace (space, tab, CR, LF) — dropped.
		if (code === 32 || code === 9 || code === 13 || code === 10) {
			i++;
			continue;
		}

		// String literal.
		if (code === 34) {
			const { value, end } = matchQuoted(source, i, 34);
			tokens.push({ type: "str", value, start: i, end });
			i = end;
			continue;
		}

		// Quoted sheet name.
		if (code === 39) {
			const { value, end } = matchQuoted(source, i, 39);
			tokens.push({ type: "sheet", value, start: i, end });
			i = end;
			continue;
		}

		// Error literal, or the bare spill operator `#`.
		if (code === 35) {
			const lit = ERROR_LITERALS.find((e) => source.startsWith(e, i));
			if (lit !== undefined) {
				tokens.push({ type: "err", value: lit, start: i, end: i + lit.length });
				i += lit.length;
				continue;
			}
			tokens.push({ type: "op", value: "#", start: i, end: i + 1 });
			i++;
			continue;
		}

		// Balanced bracket span.
		if (code === 91) {
			const end = matchBracket(source, i);
			tokens.push({ type: "bracket", value: source.slice(i, end), start: i, end });
			i = end;
			continue;
		}

		// Number (a leading `.` counts only when a digit follows, else `.` is a stray op).
		if (isDigit(code) || (code === 46 && i + 1 < len && isDigit(source.charCodeAt(i + 1)))) {
			const end = matchNumber(source, i);
			tokens.push({ type: "num", value: source.slice(i, end), start: i, end });
			i = end;
			continue;
		}

		// Cell address, or a name (function/defined name/sheet/boolean). Only `$` or an ASCII letter
		// can begin a cell; `_`, `\`, and Unicode starts are always names.
		if (code === 36 || isAsciiLetter(code)) {
			const cellEnd = matchCell(source, i);
			if (cellEnd !== -1) {
				tokens.push({
					type: "cell",
					value: source.slice(i, cellEnd),
					start: i,
					end: cellEnd,
				});
				i = cellEnd;
				continue;
			}
		}
		// A `$`-prefixed partial reference (`$A`, `$2`) — a whole-column/row endpoint. Emitted as a
		// name so the parser's reference paths (bare `:` range, `sheet!` core) treat it uniformly.
		if (code === 36) {
			const dollarEnd = matchDollarRef(source, i);
			if (dollarEnd !== -1) {
				tokens.push({
					type: "name",
					value: source.slice(i, dollarEnd),
					start: i,
					end: dollarEnd,
				});
				i = dollarEnd;
				continue;
			}
		}
		if (isNameStart(code)) {
			let j = i + 1;
			while (j < len && isNameChar(source.charCodeAt(j))) j++;
			tokens.push({ type: "name", value: source.slice(i, j), start: i, end: j });
			i = j;
			continue;
		}

		// Two-character comparison operators, then any single-character operator/punctuation.
		if (code === 60 || code === 62) {
			// '<' or '>'
			const two = source.slice(i, i + 2);
			if (two === "<=" || two === ">=" || two === "<>") {
				tokens.push({ type: "op", value: two, start: i, end: i + 2 });
				i += 2;
				continue;
			}
		}
		tokens.push({ type: "op", value: source[i] as string, start: i, end: i + 1 });
		i++;
	}
	tokens.push({ type: "eof", value: "", start: len, end: len });
	return tokens;
}
