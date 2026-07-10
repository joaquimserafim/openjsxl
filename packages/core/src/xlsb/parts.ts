import { RecordData, readRecords } from "../biff";
import { BUILTIN_FORMATS, isBuiltinDateId, isDateFormatCode } from "../ooxml";
import type { SheetState } from "../types";
import * as R from "./records";

// Parsers for the small BIFF12 parts every sheet depends on — the workbook (sheet list), the shared
// string table, and the styles table (for date detection + number-format codes). Each walks the
// record stream once; unknown records are ignored.

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

/** Parse workbook.bin → the sheet list (BrtBundleSh). */
export function parseXlsbWorkbook(bytes: Uint8Array): XlsbWorkbookMeta {
	const sheets: XlsbSheetEntry[] = [];
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.SHEET) {
			const d = new RecordData(rec.data);
			const hsState = d.u32(); // 0 = visible, 1 = hidden, 2 = veryHidden
			d.u32(); // iTabID
			const relId = d.wideString() ?? "";
			const name = d.wideString() ?? "";
			const state: SheetState =
				hsState === 2 ? "veryHidden" : hsState === 1 ? "hidden" : "visible";
			sheets.push({ name, relId, state });
		} else if (rec.id === R.WORKBOOK_END) {
			break;
		}
	}
	// The 1904 date system (BrtWbProp) is not read: its flag bit couldn't be verified against an
	// independent oracle (see F7.2 notes), and reading the wrong bit would risk mis-dating the far
	// more common 1900 files. So .xlsb dates use the 1900 epoch; 1904 is a documented follow-up.
	return { sheets, date1904: false };
}

/** Parse sharedStrings.bin → the string table (BrtSSTItem text; rich runs/phonetics dropped). */
export function parseXlsbStrings(bytes: Uint8Array): string[] {
	const strings: string[] = [];
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.SI) {
			const d = new RecordData(rec.data);
			d.skip(1); // flags
			strings.push(d.wideString() ?? "");
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
 * index to its format code and reuses the SAME `isDateFormatCode`/`isBuiltinDateId` the xlsx reader
 * uses, so date detection is identical across formats.
 */
export function parseXlsbStyles(bytes: Uint8Array): XlsbStyleTable {
	const customFormats = new Map<number, string>();
	const xfNumFmt: number[] = []; // cellXfs index → numFmtId
	let inCellXfs = false;
	for (const rec of readRecords(bytes)) {
		if (rec.id === R.FMT) {
			const d = new RecordData(rec.data);
			const ifmt = d.u16();
			const code = d.wideString();
			if (code !== undefined) customFormats.set(ifmt, code);
		} else if (rec.id === R.CELLXFS) {
			inCellXfs = true;
		} else if (rec.id === R.CELLXFS_END) {
			inCellXfs = false;
		} else if (rec.id === R.XF && inCellXfs) {
			const d = new RecordData(rec.data);
			d.skip(2); // ixfeParent
			xfNumFmt.push(d.u16()); // iFmt
		}
	}
	const numFmtId = (styleRef: number | undefined): number | undefined => xfNumFmt[styleRef ?? 0];
	return {
		isDateStyle(styleRef) {
			const id = numFmtId(styleRef);
			if (id === undefined) return false;
			const custom = customFormats.get(id);
			return custom !== undefined ? isDateFormatCode(custom) : isBuiltinDateId(id);
		},
		formatCode(styleRef) {
			const id = numFmtId(styleRef);
			if (id === undefined) return undefined;
			return customFormats.get(id) ?? BUILTIN_FORMATS[id];
		},
	};
}
