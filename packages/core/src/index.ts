// Public API surface of @openjsxl/core. Re-exports grow as features land
// (see IMPLEMENTATION.md). For now this is the cell/sheet data model.

export { XlsxError, type XlsxErrorCode } from './errors'
// Addressing + date helpers (F0.3 / F0.4).
export type { CellRef } from './ooxml/a1'
export { columnToIndex, formatRef, indexToColumn, parseRef } from './ooxml/a1'
export { dateToSerial, serialToDate } from './ooxml/dates'
// Reader API (F1.7) — open a workbook and read typed cells.
export { openXlsx, type ReadOptions, streamSheetRows, Workbook, Worksheet } from './reader/workbook'
export type { Row } from './reader/worksheet'
export type { Cell, CellType, Comment, Hyperlink, SheetInfo } from './types'
// Writer API (F3.2) — serialize a workbook described as plain data to .xlsx bytes.
export {
	type CellValue,
	type SheetInput,
	type WorkbookInput,
	type WriteOptions,
	writeXlsx,
} from './writer'
