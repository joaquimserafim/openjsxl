// writer layer — build the OPC/ZIP container (F3.1) and the OOXML workbook inside it (F3.2).
// The container primitives (writeZip, crc32, deflateRaw) stay internal; writeXlsx is the public
// surface, re-exported from the package index.

export { crc32 } from "./crc32"
export { deflateRaw } from "./deflate"
export { workbookToInput } from "./from-workbook"
export type {
	CellInput,
	CellValue,
	SheetInput,
	StyledCell,
	WorkbookInput,
	WriteOptions,
} from "./types"
export { writeXlsx } from "./workbook"
export { writeZip, type ZipInput } from "./zip"
