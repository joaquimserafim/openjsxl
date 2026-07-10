import { RecordData, readRecords } from "../biff";
import { formatRef, MAX_COL, MAX_ROW, serialToDate } from "../ooxml";
import type { Cell } from "../types";
import type { XlsbStyleTable } from "./parts";
import * as R from "./records";

// Parse one worksheet part (sheetN.bin) into the shared cell model. BIFF12 sheet data is a flat run
// of records: a ROW record sets the current 0-based row, then each cell record carries its 0-based
// column, a style index, and a typed value. Rows/columns convert to the 1-based public A1 model and
// clamp to the shared grid bounds (a corrupt record can't place a cell off-grid or throw).

/** A hyperlink as it appears in sheetN.bin: a covered range plus the rel id of its target. */
interface XlsbHyperlinkRef {
	readonly ref: string;
	readonly relId: string;
}

export interface XlsbSheetData {
	readonly cells: Map<string, Cell>;
	/** A1 ref → the cell's style index, for `numberFormat(ref)`. */
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

export function parseXlsbSheet(
	bytes: Uint8Array,
	sharedStrings: readonly string[],
	styles: XlsbStyleTable | undefined,
	date1904: boolean,
): XlsbSheetData {
	const cells = new Map<string, Cell>();
	const cellStyles = new Map<string, number>();
	const hyperlinks: XlsbHyperlinkRef[] = [];
	let dimension: string | undefined;
	let row = 0; // current 0-based row, set by ROW records

	for (const rec of readRecords(bytes)) {
		switch (rec.id) {
			case R.ROW: {
				row = new RecordData(rec.data).u32();
				break;
			}
			case R.DIMENSION: {
				const d = new RecordData(rec.data);
				const r1 = d.u32();
				const r2 = d.u32();
				const c1 = d.u32();
				const c2 = d.u32();
				dimension = rangeOf(r1, c1, r2, c2);
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
			case R.STRING:
			case R.FORMULA_STRING:
			case R.FORMULA_FLOAT:
			case R.FORMULA_BOOL:
			case R.FORMULA_BOOLERR: {
				const d = new RecordData(rec.data);
				const col = d.u32();
				// The cell's style field packs iStyleRef into the low 24 bits; bit 24 is fPhShow and
				// bits 25–31 are reserved (MS-XLSB §2.5.9). Mask to 24 bits, or an fPhShow=1 cell (CJK
				// phonetic workbooks) would carry a corrupted style index and lose date detection.
				const styleRef = d.u32() & 0xffffff;
				const ref = refOf(row, col);
				cellStyles.set(ref, styleRef);
				const cell = decodeCell(rec.id, ref, d, styleRef, sharedStrings, styles, date1904);
				if (cell !== undefined) cells.set(ref, cell);
				break;
			}
		}
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
			return numericCell(ref, d.rk(), styleRef, styles, date1904);
		case R.FLOAT:
		case R.FORMULA_FLOAT:
			return numericCell(ref, d.f64(), styleRef, styles, date1904);
		case R.STRING: {
			const idx = d.u32();
			return { ref, type: "string", value: sharedStrings[idx] ?? "" };
		}
		case R.FORMULA_STRING:
			return { ref, type: "string", value: d.wideString() ?? "" };
		case R.BOOL:
		case R.FORMULA_BOOL:
			return { ref, type: "boolean", value: d.u8() !== 0 };
		case R.BOOLERR:
		case R.FORMULA_BOOLERR:
			return { ref, type: "error", value: R.BIFF_ERRORS[d.u8()] ?? "#ERR" };
		default:
			return undefined; // BLANK: a styled empty cell — record the style, emit no value
	}
}
