// A1 reference helpers. Excel columns use bijective base-26 (A=1 … Z=26, AA=27),
// which is not ordinary base-26 because there is no zero digit.
//
// These four helpers are public API, so they honour the package-wide error contract: a bad ref
// raises `XlsxError("invalid-input")`, never a bare Error. Tolerant callers that catch and fall
// back (the reader's safeColumn, the table/range shields) are unaffected — XlsxError extends Error.

import { XlsxError } from "../errors";

export interface CellRef {
	/** 1-based column index (A = 1). */
	readonly col: number;
	/** 1-based row index. */
	readonly row: number;
}

// Excel's grid limits (XFD1048576). The tolerant READER keeps cells beyond them, faithful to the
// bytes; the WRITER and the bridge refuse them — a ref outside the grid is unopenable in Excel,
// and an absurd row number would otherwise become the length of an array someone iterates.
/** Highest row Excel supports (1,048,576). */
export const MAX_ROW = 1_048_576;
/** Highest column Excel supports (16,384 = XFD). */
export const MAX_COL = 16_384;
/** Excel's column-width ceiling (in characters of the default font). */
export const MAX_COL_WIDTH = 255;
/** Excel's row-height ceiling (in points). */
export const MAX_ROW_HEIGHT = 409.5;

const CODE_UPPER_A = 65;
const CODE_UPPER_Z = 90;
const CODE_LOWER_A = 97;
const CODE_LOWER_Z = 122;

export function columnToIndex(letters: string): number {
	if (letters.length === 0) throw new XlsxError("invalid-input", "empty column reference");
	let index = 0;
	for (let i = 0; i < letters.length; i++) {
		const code = letters.charCodeAt(i);
		let value = 0;
		if (code >= CODE_UPPER_A && code <= CODE_UPPER_Z) value = code - CODE_UPPER_A + 1;
		else if (code >= CODE_LOWER_A && code <= CODE_LOWER_Z) value = code - CODE_LOWER_A + 1;
		else throw new XlsxError("invalid-input", `invalid column reference: ${letters}`);
		index = index * 26 + value;
		// An absurdly long ref (far beyond Excel's XFD/16384 limit) overflows past exact
		// integer precision — bail before it silently becomes a lossy float or Infinity, which
		// would poison downstream column arithmetic and formatRef. Bailing here also caps work
		// on a megabyte-long attacker-supplied ref. Callers that tolerate bad refs (the reader's
		// safeColumn) catch this and fall back to positional addressing.
		if (index > Number.MAX_SAFE_INTEGER)
			throw new XlsxError("invalid-input", `column reference too large: ${letters}`);
	}
	return index;
}

export function indexToColumn(index: number): string {
	if (!Number.isInteger(index) || index < 1)
		throw new XlsxError("invalid-input", `invalid column index: ${index}`);
	let remaining = index;
	let letters = "";
	while (remaining > 0) {
		const digit = (remaining - 1) % 26;
		letters = String.fromCharCode(CODE_UPPER_A + digit) + letters;
		remaining = Math.floor((remaining - 1) / 26);
	}
	return letters;
}

const A1_PATTERN = /^([A-Za-z]+)([1-9][0-9]*)$/;

export function parseRef(ref: string): CellRef {
	const match = A1_PATTERN.exec(ref);
	if (match === null) throw new XlsxError("invalid-input", `invalid A1 reference: ${ref}`);
	return {
		col: columnToIndex(match[1] as string),
		row: Number.parseInt(match[2] as string, 10),
	};
}

export function formatRef(ref: CellRef): string {
	// Read each side of the caller's ref once: a getter must not be able to return one row to the
	// bounds check and another to the emitted string.
	const col = ref.col;
	const row = ref.row;
	if (!Number.isInteger(row) || row < 1)
		throw new XlsxError("invalid-input", `invalid row index: ${row}`);
	return `${indexToColumn(col)}${row}`;
}

// Strictly canonical A1: uppercase letters, no leading zeros — the form every real producer writes and
// the only form the writer emits. Three letters cap the column at ZZZ (18,278), so columnToIndex can't
// overflow; the grid bound is checked after.
const CANONICAL_CELL = /^[A-Z]{1,3}[1-9][0-9]*$/;

/**
 * Parse a canonical, in-grid single cell like `"A1"` to its `{col,row}`, or `undefined` when it is not
 * canonical (lowercase, leading zero, non-A1) or falls outside Excel's grid. Shared so the tolerant
 * reader (which DROPS a hostile ref) and the strict writer (which REJECTS one) agree on "valid cell".
 */
export function parseCanonicalCell(ref: string): CellRef | undefined {
	if (!CANONICAL_CELL.test(ref)) return undefined;
	const parsed = parseRef(ref);
	return parsed.col <= MAX_COL && parsed.row <= MAX_ROW ? parsed : undefined;
}

/**
 * Parse a canonical, in-grid A1 range like `"A1:C3"` (or a single cell `"A1"`) to its two corners, or
 * `undefined` when either end is not canonical / in-grid. Corners are returned as written, not sorted.
 */
export function parseCanonicalRange(ref: string): { from: CellRef; to: CellRef } | undefined {
	const colon = ref.indexOf(":");
	const from = parseCanonicalCell(colon === -1 ? ref : ref.slice(0, colon));
	const to = colon === -1 ? from : parseCanonicalCell(ref.slice(colon + 1));
	if (from === undefined || to === undefined) return undefined;
	return { from, to };
}

/**
 * True when a parsed range runs top-left → bottom-right — the only orientation Excel writes and the
 * only one the writer accepts. Single-sources that bound so the tolerant reader DROPS a backwards
 * range and the strict writer REJECTS one by the SAME rule (`parseCanonicalRange` returns corners as
 * written, so orientation is checked separately).
 */
export function rangeRunsForward(range: { from: CellRef; to: CellRef }): boolean {
	return range.to.col >= range.from.col && range.to.row >= range.from.row;
}
