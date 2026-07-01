// Public API surface of @openjsxl/core. Re-exports grow as features land
// (see IMPLEMENTATION.md). For now this is the cell/sheet data model.

export { XlsxError, type XlsxErrorCode } from './errors'
// Addressing + date helpers (F0.3 / F0.4).
export type { CellRef } from './ooxml/a1'
export { columnToIndex, formatRef, indexToColumn, parseRef } from './ooxml/a1'
export { serialToDate } from './ooxml/dates'
// Reader API (F1.7) — open a workbook and read typed cells.
export { openXlsx, streamSheetRows, Workbook, Worksheet } from './reader/workbook'
export type { Row } from './reader/worksheet'
export type { Cell, CellType, Comment, Hyperlink, SheetInfo } from './types'
