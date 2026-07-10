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
export const STRING = 0x0007; // BrtCellIsst — a shared-string index
export const FORMULA_STRING = 0x0008; // BrtFmlaString — cached string result
export const FORMULA_FLOAT = 0x0009; // BrtFmlaNum — cached number result
export const FORMULA_BOOL = 0x000a; // BrtFmlaBool
export const FORMULA_BOOLERR = 0x000b; // BrtFmlaError
export const DIMENSION = 0x0194;
export const HYPERLINK = 0x03ee;
export const SHEETDATA = 0x0191;
export const SHEETDATA_END = 0x0192;

// Workbook records.
export const SHEET = 0x019c; // BrtBundleSh — one sheet's state/id/rel/name
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
