// writer layer — build the OPC/ZIP container (F3.1) and the OOXML workbook inside it (F3.2).
// The container primitives (writeZip, crc32, deflateRaw) stay internal; the public surface
// re-exported from the package index is writeXlsx, streamXlsx, the bridge (workbookToInput), and
// the input types.

export { crc32 } from "./crc32";
export { deflateRaw } from "./deflate";
export { workbookToInput } from "./from-workbook";
export { streamXlsx } from "./stream";
export type {
	CellInput,
	CellValue,
	SheetInput,
	StreamRows,
	StreamSheetInput,
	StreamWorkbookInput,
	StyledCell,
	WorkbookInput,
	WriteOptions,
} from "./types";
export { writeXlsx } from "./workbook";
export { writeZip, type ZipInput } from "./zip";
