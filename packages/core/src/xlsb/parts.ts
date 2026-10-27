import { RecordData, readRecords } from "../biff";
import { XlsxError } from "../errors";
import {
	dateFormatDecider,
	isDateFormatCode,
	MAX_NUMBER_FORMAT_LEN,
	resolveNumberFormat,
} from "../ooxml";
import type { SheetState } from "../types";
import * as R from "./records";

// Parsers for the small BIFF12 parts every sheet depends on — the workbook (sheet list), the shared
// string table, and the styles table (for date detection + number-format codes). Each walks the
// record stream once; unknown records are ignored. Each list they build is capped below V8's array
// limit (F11.59).

/**
 * The most entries a workbook's sheet list, shared-string table or cell-format table keeps. It sits
 * below V8's array limit (2^27 elements), so a larger table fails with `XlsxError('part-too-large')`
 * instead of a bare `RangeError` (F11.59).
 */
export const MAX_XLSB_TABLE_ENTRIES = 16_000_000;

/**
 * `list.push(value)`, unless `list` already holds `maxEntries` entries: then it fails with
 * `XlsxError('part-too-large')`, naming the table as `what` (F11.59).
 *
 * @example pushWithinBudget([], "a", 1, "shared strings") // the list now holds "a"
 * @example pushWithinBudget(["a"], "b", 1, "shared strings") // throws XlsxError part-too-large
 */
export function pushWithinBudget<T>(list: T[], value: T, maxEntries: number, what: string): void {
	if (list.length >= maxEntries) {
		throw new XlsxError(
			"part-too-large",
			`xlsb: the workbook has more than ${maxEntries} ${what}, the table budget`,
		);
	}
	list.push(value);
}

/** One sheet as named in workbook.bin: its tab name, visibility, and the rel id of its part. */
export interface XlsbSheetEntry {
	readonly name: string;
	readonly relId: string;
	readonly state: SheetState;
}

export interface XlsbWorkbookMeta {
	readonly sheets: readonly XlsbSheetEntry[];
	readonly date1904: boolean;
}

/**
 * Parse workbook.bin → the sheet list (BrtBundleSh) and the date system (BrtWbProp.f1904). More than
 * `maxEntries` sheets fail with `XlsxError('part-too-large')` (F11.59).
 *
 * @example parseXlsbWorkbook(workbookBin) // { sheets: [{ name: "Sheet1", relId: "rId1", state: "visible" }], date1904: false }
 * @example parseXlsbWorkbook(workbookBin, 0) // throws XlsxError part-too-large on the first sheet
 */
export function parseXlsbWorkbook(
	bytes: Uint8Array,
	maxEntries: number = MAX_XLSB_TABLE_ENTRIES,
): XlsbWorkbookMeta {
	const sheets: XlsbSheetEntry[] = [];
	let date1904 = false;
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.WB_PROP) {
			// f1904 is bit 0 of the leading flags word; every other bit (reserved, fFilterPrivacy,
			// grbitUpdateLinks, …) is unrelated and must not enable it. A record too short to carry
			// the whole word reads 0 (the cursor's degrade contract) → the 1900 default. The record
			// precedes the sheet list in Excel's output, but the scan is order-independent.
			const flags = new RecordData(rec.data).u32();
			date1904 = (flags & 0x01) === 1;
		} else if (rec.id === R.SHEET) {
			const d = new RecordData(rec.data);
			const hsState = d.u32(); // 0 = visible, 1 = hidden, 2 = veryHidden
			d.u32(); // iTabID
			const relId = d.wideString() ?? "";
			const name = d.wideString() ?? "";
			const state: SheetState =
				hsState === 2 ? "veryHidden" : hsState === 1 ? "hidden" : "visible";
			pushWithinBudget(sheets, { name, relId, state }, maxEntries, "sheets");
		} else if (rec.id === R.WORKBOOK_END) {
			break;
		}
	}
	return { sheets, date1904 };
}

/**
 * Parse sharedStrings.bin → the string table (BrtSSTItem text; rich runs/phonetics dropped). More than
 * `maxEntries` strings fail with `XlsxError('part-too-large')` (F11.59).
 *
 * @example parseXlsbStrings(sharedStringsBin) // ["Name", "Qty"]
 * @example parseXlsbStrings(sharedStringsBin, 1) // throws XlsxError part-too-large on the second string
 */
export function parseXlsbStrings(
	bytes: Uint8Array,
	maxEntries: number = MAX_XLSB_TABLE_ENTRIES,
): string[] {
	const strings: string[] = [];
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.SI) {
			const d = new RecordData(rec.data);
			d.skip(1); // flags
			pushWithinBudget(strings, d.wideString() ?? "", maxEntries, "shared strings");
		} else if (rec.id === R.SST_END) {
			break;
		}
	}
	return strings;
}

/** Date detection + number-format lookup by a cell's style index — the same contract as the xlsx StyleTable. */
export interface XlsbStyleTable {
	isDateStyle(styleRef: number | undefined): boolean;
	formatCode(styleRef: number | undefined): string | undefined;
}

/**
 * Parse styles.bin → a style table. It collects custom number formats (BrtFmt, id ≥ 164) and each
 * cellXfs entry's numFmtId (BrtXF iFmt, at offset 2 after ixfeParent), then resolves a cell's style
 * index to its format code and reuses the SAME cached date decision and 255-character format limit
 * the xlsx reader uses (F11.50), so date detection is identical across formats. `detect` defaults
 * to `isDateFormatCode` and exists so tests can count decisions (one per id). More than `maxEntries`
 * cell formats fail with `XlsxError('part-too-large')` (F11.59).
 *
 * @example parseXlsbStyles(stylesBin).formatCode(1) // the code xf 1 uses, or undefined
 * @example parseXlsbStyles(stylesBin, isDateFormatCode, 1) // throws XlsxError part-too-large on the second cell format
 */
export function parseXlsbStyles(
	bytes: Uint8Array,
	detect: (formatCode: string) => boolean = isDateFormatCode,
	maxEntries: number = MAX_XLSB_TABLE_ENTRIES,
): XlsbStyleTable {
	const customFormats = new Map<number, string | null>();
	const xfNumFmt: number[] = []; // cellXfs index → numFmtId
	let inCellXfs = false;
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.FMT) {
			const d = new RecordData(rec.data);
			const ifmt = d.u16();
			const code = d.wideString();
			if (code !== undefined) {
				customFormats.set(ifmt, code.length <= MAX_NUMBER_FORMAT_LEN ? code : null);
			}
		} else if (rec.id === R.CELLXFS) {
			inCellXfs = true;
		} else if (rec.id === R.CELLXFS_END) {
			inCellXfs = false;
		} else if (rec.id === R.XF && inCellXfs) {
			const d = new RecordData(rec.data);
			d.skip(2); // ixfeParent
			pushWithinBudget(xfNumFmt, d.u16(), maxEntries, "cell formats"); // iFmt
		}
	}
	const numFmtId = (styleRef: number | undefined): number | undefined => xfNumFmt[styleRef ?? 0];
	const isDateFormatId = dateFormatDecider(customFormats, detect);
	return {
		isDateStyle(styleRef) {
			const id = numFmtId(styleRef);
			if (id === undefined) return false;
			return isDateFormatId(id);
		},
		formatCode(styleRef) {
			const id = numFmtId(styleRef);
			if (id === undefined) return undefined;
			return resolveNumberFormat(customFormats, id);
		},
	};
}
