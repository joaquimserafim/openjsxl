// Public API surface of @openjsxl/core. Re-exports grow as features land
// (see IMPLEMENTATION.md). For now this is the cell/sheet data model.

export { XlsxError, type XlsxErrorCode } from "./errors"
// Addressing + date helpers (F0.3 / F0.4).
export type { CellRef } from "./ooxml/a1"
export { columnToIndex, formatRef, indexToColumn, parseRef } from "./ooxml/a1"
export { dateToSerial, serialToDate } from "./ooxml/dates"
// Reader API (F1.7) — open a workbook and read typed cells.
export { openXlsx, type ReadOptions, streamSheetRows, Workbook, Worksheet } from "./reader/workbook"
export type { Row } from "./reader/worksheet"
// Style model (F4.1) — shared by the reader (`Worksheet.style(ref)`) and, from F4.2, the writer.
export type {
	Alignment,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	Cell,
	CellStyle,
	CellType,
	Color,
	ColumnProps,
	Comment,
	FillStyle,
	FontStyle,
	FreezePane,
	HorizontalAlignment,
	Hyperlink,
	PatternType,
	RowProps,
	SheetInfo,
	SheetState,
	UnderlineStyle,
	VerticalAlignment,
} from "./types"
// Writer API (F3.2) — serialize a workbook described as plain data to .xlsx bytes; styled cells
// from F4.2. The bridge (F3.3) turns an open Workbook back into writer input, closing the round trip.
export {
	type CellInput,
	type CellValue,
	type SheetInput,
	type StyledCell,
	type WorkbookInput,
	type WriteOptions,
	workbookToInput,
	writeXlsx,
} from "./writer"
