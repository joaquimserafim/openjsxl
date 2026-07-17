// Public API surface of @openjsxl/core: the reader + shared Workbook/Worksheet model, the writer and
// its bridge, the style/geometry/metadata types, and the addressing/date/detection helpers. This is
// the frozen 1.0 surface (see IMPLEMENTATION.md) — everything exported here is public and documented;
// nothing else in the package is API. Every public function reports failure by throwing `XlsxError`
// (branch on its `code`), never a bare `Error`.

export { XlsxError, type XlsxErrorCode } from "./errors";
// Addressing + date helpers (F0.3 / F0.4).
export type { CellRef } from "./ooxml/a1";
export { columnToIndex, formatRef, indexToColumn, parseRef } from "./ooxml/a1";
export { dateToSerial, serialToDate } from "./ooxml/dates";
// Defined (named) ranges/constants read from the workbook (F8.2 — consumed by openjsxl/formula).
export type { DefinedName } from "./ooxml/workbook";
// CSV/TSV reader (F7.3) — open delimited text and read the SAME public Workbook surface.
export { type CsvReadOptions, openCsv } from "./reader/csv";
// Format detection (F7.4) — sniff bytes to route to the right open* function.
export { detectSpreadsheetFormat, type SpreadsheetFormat } from "./reader/detect";
// ODS reader (F7.1) — open an .ods and read the SAME public Workbook surface.
export { openOds } from "./reader/ods";
// Reader API (F1.7) — open a workbook and read typed cells.
export {
	openXlsx,
	type ReadOptions,
	streamSheetRows,
	Workbook,
} from "./reader/workbook";
// XLSB reader (F7.2) — open an Excel Binary Workbook and read the SAME public Workbook surface.
export { openXlsb } from "./reader/xlsb";
// Style model (F4.1) — shared by the reader (`Worksheet.style(ref)`) and, from F4.2, the writer.
export type {
	Alignment,
	AnchorPoint,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	Cell,
	CellProtection,
	CellStyle,
	CellType,
	CfColorScaleRule,
	CfDataBarRule,
	CfHighlightRule,
	CfHighlightType,
	CfIconSetRule,
	Cfvo,
	Color,
	ColumnProps,
	Comment,
	ConditionalFormatting,
	ConditionalFormattingRule,
	DataValidation,
	DataValidationErrorStyle,
	DataValidationOperator,
	DataValidationType,
	DxfFill,
	DxfStyle,
	FillStyle,
	FontStyle,
	FreezePane,
	HeaderFooter,
	HorizontalAlignment,
	Hyperlink,
	ImageAnchor,
	PageMargins,
	PageSetup,
	PatternType,
	PrintOptions,
	Row,
	RowProps,
	SheetAutoFilter,
	SheetImage,
	SheetInfo,
	SheetProtection,
	SheetState,
	TableColumn,
	TableInfo,
	TableStyleInfo,
	UnderlineStyle,
	VerticalAlignment,
	WorkbookProtection,
	Worksheet,
} from "./types";
// Writer API (F3.2) — serialize a workbook described as plain data to .xlsx bytes; styled cells
// from F4.2. The bridge (F3.3) turns an open Workbook back into writer input, closing the round trip.
export {
	type CellInput,
	type CellValue,
	type SheetInput,
	type StreamRows,
	type StreamSheetInput,
	type StreamWorkbookInput,
	type StyledCell,
	streamXlsx,
	type WorkbookInput,
	type WriteOptions,
	workbookToInput,
	writeXlsx,
} from "./writer";
