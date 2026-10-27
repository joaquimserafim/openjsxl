// xlsb layer — parse the BIFF12 binary parts of an .xlsb into the shared cell model (M7).
export {
	MAX_XLSB_TABLE_ENTRIES,
	parseXlsbStrings,
	parseXlsbStyles,
	parseXlsbWorkbook,
	pushWithinBudget,
	type XlsbSheetEntry,
	type XlsbStyleTable,
	type XlsbWorkbookMeta,
} from "./parts";
export {
	MAX_XLSB_CELLS,
	parseXlsbSheet,
	setWithinBudget,
	type XlsbSheetData,
} from "./sheet";
