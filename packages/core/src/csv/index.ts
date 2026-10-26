// csv layer — parse delimited text (.csv / .tsv) into the shared cell model (M7).
export {
	type CsvCellData,
	collapseQuotes,
	inferCsvValue,
	MAX_CSV_CELLS,
	parseDelimited,
	scanDelimited,
	sniffDelimiter,
} from "./parse";
