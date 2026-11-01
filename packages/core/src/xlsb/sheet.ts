import { RecordData, readRecords } from "../biff";
import { XlsxError } from "../errors";
import { formatRef, MAX_COL, MAX_ROW, serialToDate } from "../ooxml";
import type { Cell } from "../types";
import type { XlsbStyleTable } from "./parts";
import * as R from "./records";

// Parse one worksheet part (sheetN.bin) into the shared cell model. BIFF12 sheet data is a flat run
// of records: a ROW record sets the current 0-based row; full cells carry their 0-based column,
// while short cells use the previous column + 1. Both carry a style index and a typed value.
// Rows/columns convert to the 1-based public A1 model and clamp to the shared grid bounds
// (a corrupt record can't place a cell off-grid). Style 0, the default, is not stored, and the cells
// and stored styles of one sheet are capped below V8's Map limit (F11.59). A record cut before the end
// of its value yields no value (F11.100).

/**
 * The most cells, and the most cells with a non-default style, one xlsb sheet keeps. It sits below
 * V8's Map limit (2^24 entries), so a larger sheet fails with `XlsxError('part-too-large')` instead of
 * a bare `RangeError` (F11.59).
 */
export const MAX_XLSB_CELLS = 16_000_000;

/** A hyperlink as it appears in sheetN.bin: a covered range plus the rel id of its target. */
interface XlsbHyperlinkRef {
	readonly ref: string;
	readonly relId: string;
}

export interface XlsbSheetData {
	readonly cells: Map<string, Cell>;
	/**
	 * A1 ref → the cell's style index, for `numberFormat(ref)`. Style 0 is not stored: a lookup that
	 * finds nothing falls back to it.
	 */
	readonly cellStyles: Map<string, number>;
	readonly dimension: string | undefined;
	readonly hyperlinks: readonly XlsbHyperlinkRef[];
}

// A1 ref for a 0-based (row, col), clamped into the grid (reader degrades; never throws).
function refOf(row0: number, col0: number): string {
	const col = Math.min(col0, MAX_COL - 1) + 1;
	const row = Math.min(row0, MAX_ROW - 1) + 1;
	return formatRef({ col, row });
}

function rangeOf(r1: number, c1: number, r2: number, c2: number): string {
	const tl = refOf(r1, c1);
	const br = refOf(r2, c2);
	return tl === br ? tl : `${tl}:${br}`;
}

/**
 * Decode a numeric cell (RK, real, or a cached numeric formula result), applying date detection via
 * the style table — mirroring the xlsx value path (F2.1/F3.3): a date-styled number becomes a `Date`
 * unless the serial falls outside JS's Date range, in which case it stays a number.
 */
function numericCell(
	ref: string,
	value: number,
	styleRef: number,
	styles: XlsbStyleTable | undefined,
	date1904: boolean,
): Cell {
	if (styles?.isDateStyle(styleRef)) {
		const date = serialToDate(value, date1904);
		if (!Number.isNaN(date.getTime())) return { ref, type: "date", value: date };
	}
	return { ref, type: "number", value };
}

/**
 * `map.set(ref, value)`, unless `ref` is new and `map` already holds `maxCells` entries: then it fails
 * with `XlsxError('part-too-large')`, naming the entries as `what`. Replacing a ref the map holds is
 * always allowed (F11.59).
 *
 * @example setWithinBudget(new Map(), "A1", 1, 1, "cells") // the map now holds A1
 * @example setWithinBudget(new Map([["A1", 1]]), "B1", 2, 1, "cells") // throws XlsxError part-too-large
 */
export function setWithinBudget<V>(
	map: Map<string, V>,
	ref: string,
	value: V,
	maxCells: number,
	what: string,
): void {
	if (map.size >= maxCells && !map.has(ref)) {
		throw new XlsxError(
			"part-too-large",
			`xlsb: a sheet has more than ${maxCells} ${what}, the cell budget`,
		);
	}
	map.set(ref, value);
}

/**
 * The bytes a cell record's value takes after its `header` (4 bytes for a short form, 8 for a full one),
 * as `decodeCell` reads them: 4 for an RK number or a shared-string index, 8 for a real, 1 for a Boolean
 * or an error, and none for a blank. A wide string takes its 4-byte count plus its UTF-16 code units;
 * the `0xFFFFFFFF` null sentinel, or a count the record cuts, takes 4 (F11.100).
 *
 * @example valueLength(R.BLANK, data, 8) // 0
 * @example valueLength(R.SHORT_NUM, data, 4) // 4
 * @example valueLength(R.FLOAT, data, 8) // 8
 * @example valueLength(R.BOOL, data, 8) // 1
 * @example valueLength(R.SHORT_INLINE_STRING, Uint8Array.of(0, 0, 0, 0, 3, 0, 0, 0), 4) // 10
 * @example valueLength(R.SHORT_INLINE_STRING, Uint8Array.of(0, 0, 0, 0, 255, 255, 255, 255), 4) // 4
 * @example valueLength(R.INLINE_STRING, Uint8Array.of(0, 0, 0, 0), 8) // 4: the count is cut
 */
export function valueLength(id: number, data: Uint8Array, header: number): number {
	switch (id) {
		case R.NUM:
		case R.SHORT_NUM:
		case R.STRING:
		case R.SHORT_STRING:
			return 4;
		case R.FLOAT:
		case R.SHORT_FLOAT:
		case R.FORMULA_FLOAT:
			return 8;
		case R.BOOL:
		case R.SHORT_BOOL:
		case R.FORMULA_BOOL:
		case R.BOOLERR:
		case R.SHORT_BOOLERR:
		case R.FORMULA_BOOLERR:
			return 1;
		case R.INLINE_STRING:
		case R.SHORT_INLINE_STRING:
		case R.FORMULA_STRING: {
			if (data.length < header + 4) return 4; // the record cuts the count
			// The little-endian u32 count, read in place: this runs for every string record.
			let count = 0;
			for (let i = 3; i >= 0; i--) count = count * 256 + (data[header + i] as number);
			return count === 0xffffffff ? 4 : 4 + count * 2;
		}
		default:
			return 0; // Full/short BLANK: no value
	}
}

/**
 * Parse one worksheet part into typed cells, their non-default styles, the dimension and hyperlink
 * refs. A sheet with more than `maxCells` cells, or more than `maxCells` cells with a non-default
 * style, fails with `XlsxError('part-too-large')` (F11.59).
 *
 * @example parseXlsbSheet(sheetBin, sharedStrings, styles, false) // { cells, cellStyles, dimension, hyperlinks }
 * @example parseXlsbSheet(sheetBin, [], undefined, false, 1) // throws XlsxError part-too-large past one cell
 */
export function parseXlsbSheet(
	bytes: Uint8Array,
	sharedStrings: readonly string[],
	styles: XlsbStyleTable | undefined,
	date1904: boolean,
	maxCells: number = MAX_XLSB_CELLS,
): XlsbSheetData {
	const cells = new Map<string, Cell>();
	const cellStyles = new Map<string, number>();
	const hyperlinks: XlsbHyperlinkRef[] = [];
	let dimension: string | undefined;
	let emptyDimension = false;
	let firstUsedRow = Infinity;
	let firstUsedCol = Infinity;
	let lastUsedRow = 0;
	let lastUsedCol = 0;
	let row = 0; // current 0-based row, set by ROW records
	let col = -1; // raw previous column; clamp only when converting to a public ref

	for (const rec of readRecords(bytes)) {
		switch (rec.id) {
			case R.ROW: {
				const d = new RecordData(rec.data);
				row = d.u32();
				col = -1;
				d.skip(7); // ixfe, height, first flags byte; fGhostDirty is bit 6 of byte 11
				if ((d.u8() & 0x40) !== 0) {
					firstUsedRow = Math.min(firstUsedRow, row);
					lastUsedRow = Math.max(lastUsedRow, row);
				}
				break;
			}
			case R.COL_INFO: {
				// A column default counts only with a non-default XF (F11.37). XF 0 is the default
				// format, so a width-only or hidden-only column is not formatting, and a producer that
				// wrote an all-zero BrtWsDim did not count it. Require the full fields (bounds, width,
				// XF) so a truncated record cannot invent formatting evidence.
				if (rec.data.length < 16) break;
				const d = new RecordData(rec.data);
				const first = d.u32();
				const last = d.u32();
				d.skip(4); // coldx, the width
				const ixfe = d.u32();
				if (first <= last && ixfe !== 0) {
					firstUsedCol = Math.min(firstUsedCol, first);
					lastUsedCol = Math.max(lastUsedCol, last);
				}
				break;
			}
			case R.DIMENSION: {
				const d = new RecordData(rec.data);
				const r1 = d.u32();
				const r2 = d.u32();
				const c1 = d.u32();
				const c2 = d.u32();
				// All-zero BrtWsDim can mean either an empty sheet or a genuine A1 used range.
				// Resolve it after reading cells/default formats; other ranges keep their clamping.
				emptyDimension = r1 === 0 && r2 === 0 && c1 === 0 && c2 === 0;
				const invalidDimension =
					r1 === 0xffffffff &&
					r2 === 0xffffffff &&
					c1 === 0xffffffff &&
					c2 === 0xffffffff;
				dimension =
					emptyDimension || invalidDimension ? undefined : rangeOf(r1, c1, r2, c2);
				break;
			}
			case R.HYPERLINK: {
				const d = new RecordData(rec.data);
				const r1 = d.u32();
				const r2 = d.u32();
				const c1 = d.u32();
				const c2 = d.u32();
				const relId = d.wideString();
				if (relId !== undefined && relId !== "") {
					hyperlinks.push({ ref: rangeOf(r1, c1, r2, c2), relId });
				}
				break;
			}
			case R.BLANK:
			case R.NUM:
			case R.BOOL:
			case R.BOOLERR:
			case R.FLOAT:
			case R.INLINE_STRING:
			case R.STRING:
			case R.FORMULA_STRING:
			case R.FORMULA_FLOAT:
			case R.FORMULA_BOOL:
			case R.FORMULA_BOOLERR:
			case R.RICH_STRING:
			case R.SHORT_BLANK:
			case R.SHORT_NUM:
			case R.SHORT_BOOLERR:
			case R.SHORT_BOOL:
			case R.SHORT_FLOAT:
			case R.SHORT_INLINE_STRING:
			case R.SHORT_STRING: {
				const d = new RecordData(rec.data);
				// Short cells omit only the column field. Every cell, even a blank or truncated
				// record, establishes the position for the next short cell in this row.
				const short = rec.id >= R.SHORT_BLANK && rec.id <= R.SHORT_STRING;
				col = short ? col + 1 : d.u32();
				// All cell records count as used, including styled blanks.
				firstUsedRow = Math.min(firstUsedRow, row);
				firstUsedCol = Math.min(firstUsedCol, col);
				lastUsedRow = Math.max(lastUsedRow, row);
				lastUsedCol = Math.max(lastUsedCol, col);
				// Recognize the full rich-string anchor so a following short cell stays in place.
				// Its RichStr payload and style remain outside the supported value model.
				if (rec.id === R.RICH_STRING) break;
				// The cell's style field packs iStyleRef into the low 24 bits; bit 24 is fPhShow and
				// bits 25–31 are reserved (MS-XLSB Cell). Mask to 24 bits, or an fPhShow=1 cell (CJK
				// phonetic workbooks) would carry a corrupted style index and lose date detection.
				const styleRef = d.u32() & 0xffffff;
				// Style 0 is what every lookup falls back to, so it is not stored, and a style-0 blank
				// needs no ref unless an earlier record styled one. A style-0 record still replaces an
				// earlier style at the same ref, as when every style was stored (F11.59).
				const blank = rec.id === R.BLANK || rec.id === R.SHORT_BLANK;
				if (blank && styleRef === 0 && cellStyles.size === 0) break;
				const ref = refOf(row, col);
				if (styleRef === 0) cellStyles.delete(ref);
				else
					setWithinBudget(
						cellStyles,
						ref,
						styleRef,
						maxCells,
						"cells with a non-default style",
					);
				// A record cut before the end of its value reads like a blank of its form: its style, no
				// value, and an earlier value at the ref survives. Past the end, RecordData would read
				// 0, `false` or a partial string instead (F11.100).
				const header = short ? 4 : 8;
				if (rec.data.length < header + valueLength(rec.id, rec.data, header)) break;
				const cell = decodeCell(rec.id, ref, d, styleRef, sharedStrings, styles, date1904);
				if (cell !== undefined) setWithinBudget(cells, ref, cell, maxCells, "cells");
				break;
			}
		}
	}

	if (emptyDimension && (firstUsedRow !== Infinity || firstUsedCol !== Infinity)) {
		// BrtWsDim counts row/column defaults only on their own axis. Supply A/1 for an
		// otherwise missing axis after the scan, so record order cannot inflate the bounds.
		dimension = rangeOf(
			firstUsedRow === Infinity ? 0 : firstUsedRow,
			firstUsedCol === Infinity ? 0 : firstUsedCol,
			lastUsedRow,
			lastUsedCol,
		);
	}

	return { cells, cellStyles, dimension, hyperlinks };
}

function decodeCell(
	id: number,
	ref: string,
	d: RecordData,
	styleRef: number,
	sharedStrings: readonly string[],
	styles: XlsbStyleTable | undefined,
	date1904: boolean,
): Cell | undefined {
	switch (id) {
		case R.NUM:
		case R.SHORT_NUM:
			return numericCell(ref, d.rk(), styleRef, styles, date1904);
		case R.FLOAT:
		case R.SHORT_FLOAT:
		case R.FORMULA_FLOAT:
			return numericCell(ref, d.f64(), styleRef, styles, date1904);
		case R.STRING:
		case R.SHORT_STRING: {
			const idx = d.u32();
			return { ref, type: "string", value: sharedStrings[idx] ?? "" };
		}
		case R.FORMULA_STRING:
		case R.INLINE_STRING:
		case R.SHORT_INLINE_STRING:
			return { ref, type: "string", value: d.wideString() ?? "" };
		case R.BOOL:
		case R.SHORT_BOOL:
		case R.FORMULA_BOOL:
			return { ref, type: "boolean", value: d.u8() !== 0 };
		case R.BOOLERR:
		case R.SHORT_BOOLERR:
		case R.FORMULA_BOOLERR:
			return { ref, type: "error", value: R.BIFF_ERRORS[d.u8()] ?? "#ERR" };
		default:
			return undefined; // Full/short BLANK: record the style, emit no value
	}
}
