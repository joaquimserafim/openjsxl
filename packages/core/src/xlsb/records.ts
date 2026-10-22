// BIFF12 record ids used by the .xlsb reader (M7). Values follow pyxlsb's byte-combined encoding:
// a constant `C` is written on the wire as `C < 0x80 ? [C] : [C & 0xff, C >> 8]`, which reproduces
// exactly the bytes Excel emits — so these decode real files and round-trip our own fixtures.

// Worksheet cell + structure records.
export const ROW = 0x0000;
export const BLANK = 0x0001;
export const NUM = 0x0002; // BrtCellRk — an RK number
export const BOOLERR = 0x0003; // BrtCellError — an error code
export const BOOL = 0x0004; // BrtCellBool
export const FLOAT = 0x0005; // BrtCellReal — an 8-byte double
export const INLINE_STRING = 0x0006; // BrtCellSt — inline UTF-16 string
export const STRING = 0x0007; // BrtCellIsst — a shared-string index
export const FORMULA_STRING = 0x0008; // BrtFmlaString — cached string result
export const FORMULA_FLOAT = 0x0009; // BrtFmlaNum — cached number result
export const FORMULA_BOOL = 0x000a; // BrtFmlaBool
export const FORMULA_BOOLERR = 0x000b; // BrtFmlaError
// Undocumented short forms emitted by SheetJS: same style/value fields, implicit next column.
export const SHORT_BLANK = 0x000c; // BrtShortBlank
export const SHORT_NUM = 0x000d; // BrtShortRk
export const SHORT_BOOLERR = 0x000e; // BrtShortError
export const SHORT_BOOL = 0x000f; // BrtShortBool
export const SHORT_FLOAT = 0x0010; // BrtShortReal
export const SHORT_INLINE_STRING = 0x0011; // BrtShortSt
export const SHORT_STRING = 0x0012; // BrtShortIsst
export const COL_INFO = 0x003c; // BrtColInfo — column bounds, width, and default format
export const RICH_STRING = 0x003e; // BrtCellRString — column anchor; rich value not decoded
export const DIMENSION = 0x0194;
export const HYPERLINK = 0x03ee;
export const SHEETDATA = 0x0191;
export const SHEETDATA_END = 0x0192;

// Workbook records.
export const SHEET = 0x019c; // BrtBundleSh — one sheet's state/id/rel/name
// BrtWbProp (MS-XLSB §2.4.866, record type 0x0099 → varint bytes 99 01): flags(u32) whose bit 0 is
// f1904 (the 1904 date system), then dwThemeVersion(u32) and strName (the workbook code name).
export const WB_PROP = 0x0199;
export const WORKBOOK_END = 0x0184;

// Shared-strings records.
export const SI = 0x0013; // BrtSSTItem
export const SST_END = 0x01a0;

// Styles records.
export const FMT = 0x002c; // BrtFmt — a custom number format (id ≥ 164)
export const XF = 0x002f; // BrtXF — ixfeParent(u16) + iFmt(u16) + …
export const CELLXFS = 0x04e9;
export const CELLXFS_END = 0x04ea;

// The BIFF error-code byte → its display text (BErr, MS-XLSB §2.5.10).
export const BIFF_ERRORS: Readonly<Record<number, string>> = {
	0: "#NULL!",
	7: "#DIV/0!",
	15: "#VALUE!",
	23: "#REF!",
	29: "#NAME?",
	36: "#NUM!",
	42: "#N/A",
	43: "#GETTING_DATA",
};
