import type { Cell } from "../types";

// Parse delimited text (.csv / .tsv) into rows of typed cells. Two pure pieces: an RFC 4180 scanner
// that splits text into raw string fields, and conservative type inference over each field. No
// container, no async — the input bytes ARE the bound; a pathological input (an unterminated quote,
// a million columns) stays linear and never hangs.

/**
 * RFC 4180 scanner: text → rows of raw string fields. A double-quote starts a quoted field, where
 * `""` is a literal quote and delimiters / newlines are ordinary text; a UTF-16 BOM is stripped;
 * CR, LF, and CRLF all end a row. An unterminated quote consumes to end-of-input as one field.
 */
export function parseDelimited(input: string, delimiter: string): string[][] {
	let text = input;
	if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // strip a leading BOM
	const rows: string[][] = [];
	let row: string[] = [];
	let field = "";
	let inQuotes = false;
	let started = false; // has the current row begun a field/quote? (distinguishes a trailing newline)
	const n = text.length;

	for (let i = 0; i < n; i++) {
		const ch = text[i];
		if (inQuotes) {
			if (ch === '"') {
				if (text[i + 1] === '"') {
					field += '"';
					i += 1; // consume the escaped quote
				} else {
					inQuotes = false;
				}
			} else {
				field += ch;
			}
			continue;
		}
		if (ch === '"' && field === "") {
			// A quote is special ONLY at the start of a field (RFC 4180); a quote in the middle of an
			// unquoted field (`a"b`) is a literal character — matching Python's csv.
			inQuotes = true;
			started = true;
		} else if (ch === delimiter) {
			row.push(field);
			field = "";
			started = true;
		} else if (ch === "\n" || ch === "\r") {
			row.push(field);
			rows.push(row);
			row = [];
			field = "";
			started = false;
			if (ch === "\r" && text[i + 1] === "\n") i += 1; // CRLF is one line ending
		} else {
			field += ch;
			started = true;
		}
	}
	// Flush the last row unless the input ended exactly on a row terminator (no trailing empty row).
	if (started || field !== "" || row.length > 0) {
		row.push(field);
		rows.push(row);
	}
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
 */
export function inferCsvValue(field: string): CsvCellData | undefined {
	if (field === "") return undefined;
	const lower = field.toLowerCase();
	if (lower === "true") return { type: "boolean", value: true };
	if (lower === "false") return { type: "boolean", value: false };
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
 * quotes. Ties and an empty first line default to comma.
 */
export function sniffDelimiter(text: string): "," | "\t" | ";" {
	const start = text.charCodeAt(0) === 0xfeff ? 1 : 0;
	let line = "";
	let inQuotes = false;
	for (let i = start; i < text.length; i++) {
		const ch = text[i];
		if (ch === '"') inQuotes = !inQuotes;
		else if (!inQuotes && (ch === "\n" || ch === "\r")) break;
		line += ch;
	}
	const unquoted = line.replace(/"[^"]*"/g, "");
	const candidates: readonly ["," | "\t" | ";", number][] = [
		[",", count(unquoted, ",")],
		["\t", count(unquoted, "\t")],
		[";", count(unquoted, ";")],
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

function count(haystack: string, needle: string): number {
	let n = 0;
	let i = haystack.indexOf(needle);
	while (i !== -1) {
		n += 1;
		i = haystack.indexOf(needle, i + 1);
	}
	return n;
}
