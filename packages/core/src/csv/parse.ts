import type { Cell } from "../types";

// Parse delimited text (.csv / .tsv) into rows of typed cells. Two pure pieces: an RFC 4180 scanner
// that splits text into raw string fields, and conservative type inference over each field. No
// container, no async — the input bytes ARE the bound; a pathological input (an unterminated quote,
// a million columns) stays linear and never hangs. The scanner walks the text by index and takes each
// field with at most two slices, so it builds no string a character at a time (F11.58).

const BOM = 0xfeff;
const QUOTE = 0x22;
const LF = 0x0a;
const CR = 0x0d;
const COMMA = 0x2c;
const TAB = 0x09;
const SEMICOLON = 0x3b;

/**
 * The most non-empty cells `openCsv` places. It sits below V8's Map limit (2^24 entries), so a
 * larger input fails with `XlsxError('part-too-large')` instead of a bare `RangeError` (F11.58).
 */
export const MAX_CSV_CELLS = 16_000_000;

// collapseQuotes joins the runs between escaped quotes this many at a time.
const JOIN_BATCH = 4096;

/**
 * `input.slice(from, end)` with each `""` collapsed to one quote, as one flat string. Every quote in
 * the range must belong to a `""` pair, as in a scanned quoted field. The runs between pairs are
 * joined in batches of 4,096, then the batches once, so no array grows past the input's length /
 * 4,096. (`replaceAll` builds a rope of about two nodes per pair, near 28 bytes of heap per
 * character, F11.58.)
 *
 * @example collapseQuotes('x"a""b""c"', 2, 9) // 'a"b"c'
 */
export function collapseQuotes(input: string, from: number, end: number): string {
	const batches: string[] = [];
	let runs: string[] = [];
	let start = from;
	for (let q = input.indexOf('""', start); q !== -1 && q < end; q = input.indexOf('""', start)) {
		runs.push(input.slice(start, q + 1)); // the run and one quote of the pair
		start = q + 2;
		if (runs.length === JOIN_BATCH) {
			batches.push(runs.join(""));
			runs = [];
		}
	}
	runs.push(input.slice(start, end));
	batches.push(runs.join(""));
	return batches.join("");
}

/**
 * RFC 4180 scanner, one field at a time: calls `onField(field, row, col)` (0-based) for every field,
 * in order. A double-quote starts a quoted field, where `""` is a literal quote and delimiters /
 * newlines are ordinary text; a UTF-16 BOM is stripped; CR, LF, and CRLF all end a row. An
 * unterminated quote consumes to end-of-input as one field. A field is at most two slices of the text,
 * its quoted part (collapsed flat by {@link collapseQuotes} when it holds `""`) and the text after it,
 * and the scanner keeps nothing between calls.
 *
 * @example scanDelimited('a,"b""c"\n1', ",", (f, r, c) => seen.push([f, r, c])) // seen: [["a", 0, 0], ['b"c', 0, 1], ["1", 1, 0]]
 */
export function scanDelimited(
	input: string,
	delimiter: string,
	onField: (field: string, row: number, col: number) => void,
): void {
	const n = input.length;
	// One UTF-16 unit. Any other delimiter never matches, as in the character-by-character scan, and a
	// plain-JS caller's non-string never throws.
	const delim =
		typeof delimiter === "string" && delimiter.length === 1 ? delimiter.charCodeAt(0) : -1;
	let i = input.charCodeAt(0) === BOM ? 1 : 0;
	let row = 0;
	let col = 0;
	let afterDelimiter = false; // a delimiter that ends the input still owes one empty field
	while (i < n) {
		let field = "";
		let p = i;
		if (input.charCodeAt(i) === QUOTE) {
			// A quote is special ONLY at the start of a field (RFC 4180). The field runs to the first quote
			// that is not followed by another; each `""` before it is one literal quote.
			let close = input.indexOf('"', i + 1);
			let escaped = false;
			while (close !== -1 && input.charCodeAt(close + 1) === QUOTE) {
				escaped = true;
				close = input.indexOf('"', close + 2);
			}
			const end = close === -1 ? n : close;
			field = escaped ? collapseQuotes(input, i + 1, end) : input.slice(i + 1, end);
			p = close === -1 ? n : close + 1;
		}
		// Unquoted text, or text after a closing quote, runs to the next delimiter or line end. A quote in
		// it is a literal character (`a"b`, `"a"b`), matching Python's csv.
		const start = p;
		let c = -1;
		while (p < n) {
			c = input.charCodeAt(p);
			if (c === delim || c === LF || c === CR) break;
			p++;
		}
		if (p > start) field += input.slice(start, p);
		onField(field, row, col);
		afterDelimiter = false;
		if (p >= n) break;
		if (c === delim) {
			col++;
			afterDelimiter = true;
			i = p + 1;
		} else {
			row++;
			col = 0;
			i = c === CR && input.charCodeAt(p + 1) === LF ? p + 2 : p + 1; // CRLF is one line ending
		}
	}
	// A trailing delimiter ends with an empty field; a trailing newline adds no empty row.
	if (afterDelimiter) onField("", row, col);
}

/**
 * RFC 4180 scanner: text → rows of raw string fields, as {@link scanDelimited} reads them. A
 * double-quote starts a quoted field, where `""` is a literal quote and delimiters / newlines are
 * ordinary text; a UTF-16 BOM is stripped; CR, LF, and CRLF all end a row. An unterminated quote
 * consumes to end-of-input as one field. It keeps every field, so use `scanDelimited` to keep fewer.
 *
 * @example parseDelimited('a,"b,c"\r\n1', ",") // [["a", "b,c"], ["1"]]
 */
export function parseDelimited(input: string, delimiter: string): string[][] {
	const rows: string[][] = [];
	scanDelimited(input, delimiter, (field, _row, col) => {
		if (col === 0) rows.push([field]);
		else (rows[rows.length - 1] as string[]).push(field);
	});
	return rows;
}

/** A field's inferred type + value, or undefined for an empty field (which reads as an empty cell). */
export interface CsvCellData {
	readonly type: Cell["type"];
	readonly value: string | number | boolean | null;
}

// A plain numeric literal: optional sign, then an integer with NO leading zeros (so "007" and ZIP
// codes stay strings), an optional fraction, and an optional exponent — or a bare fraction (".5").
// Excludes Infinity / NaN / hex (0x…) / thousands separators, which stay strings.
const NUMERIC = /^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$|^[+-]?\.\d+(?:[eE][+-]?\d+)?$/;
const PURE_INTEGER = /^[+-]?\d+$/;

/**
 * Conservative type inference for one field. `TRUE`/`FALSE` (case-insensitive) → boolean; a plain
 * numeric literal → number — EXCEPT a big-integer literal beyond `Number.MAX_SAFE_INTEGER`, which
 * stays a string to preserve its digits (IDs, account numbers). Dates are NEVER inferred (CSV date
 * formats are locale-ambiguous, so guessing fabricates wrong values). Everything else stays a string.
 *
 * @example inferCsvValue("") // undefined
 * @example inferCsvValue("True") // { type: "boolean", value: true }
 * @example inferCsvValue("FALSE") // { type: "boolean", value: false }
 * @example inferCsvValue("-2.5e3") // { type: "number", value: -2500 }
 * @example inferCsvValue("007") // { type: "string", value: "007" }
 */
export function inferCsvValue(field: string): CsvCellData | undefined {
	if (field === "") return undefined;
	// Only a 4- or 5-unit field can lowercase to "true" or "false". Lowercasing a huge field could
	// pass the maximum string length (U+0130 lowercases to two units), which kills the process (F11.58).
	if (field.length === 4 || field.length === 5) {
		const lower = field.toLowerCase();
		if (lower === "true") return { type: "boolean", value: true };
		if (lower === "false") return { type: "boolean", value: false };
	}
	if (NUMERIC.test(field)) {
		const num = Number(field);
		if (Number.isFinite(num) && !(PURE_INTEGER.test(field) && !Number.isSafeInteger(num))) {
			return { type: "number", value: num };
		}
	}
	return { type: "string", value: field };
}

/**
 * Sniff the delimiter from the first line: the most frequent of comma / tab / semicolon outside
 * quotes. Ties and an empty first line default to comma. A quote toggles quoting (no `""` escape
 * here) and a line end inside quotes does not end the line. A last, unmatched quote quotes nothing,
 * so the candidates after it still count. It counts in place, without copying the line (F11.58).
 *
 * @example sniffDelimiter('"a,b,c";d;e\nf,g,h,i') // ";"
 */
export function sniffDelimiter(text: string): "," | "\t" | ";" {
	let comma = 0;
	let tab = 0;
	let semicolon = 0;
	// Candidates since the last opening quote: dropped when it closes, kept if it never does.
	let quotedComma = 0;
	let quotedTab = 0;
	let quotedSemicolon = 0;
	let inQuotes = false;
	for (let i = text.charCodeAt(0) === BOM ? 1 : 0; i < text.length; i++) {
		const c = text.charCodeAt(i);
		if (c === QUOTE) {
			inQuotes = !inQuotes;
			quotedComma = 0;
			quotedTab = 0;
			quotedSemicolon = 0;
		} else if (!inQuotes && (c === LF || c === CR)) {
			break;
		} else if (inQuotes) {
			if (c === COMMA) quotedComma++;
			else if (c === TAB) quotedTab++;
			else if (c === SEMICOLON) quotedSemicolon++;
		} else if (c === COMMA) {
			comma++;
		} else if (c === TAB) {
			tab++;
		} else if (c === SEMICOLON) {
			semicolon++;
		}
	}
	if (inQuotes) {
		comma += quotedComma;
		tab += quotedTab;
		semicolon += quotedSemicolon;
	}
	const candidates: readonly ["," | "\t" | ";", number][] = [
		[",", comma],
		["\t", tab],
		[";", semicolon],
	];
	let best: "," | "\t" | ";" = ",";
	let bestCount = 0;
	for (const [ch, c] of candidates) {
		if (c > bestCount) {
			bestCount = c;
			best = ch;
		}
	}
	return best;
}
