// Post-M7 cross-format READ benchmark. One workload's data, authored in four container formats, each
// read by the libraries that support it. The point is format-vs-format (how fast each reader parses
// the same logical data in xlsx / xlsb / ods / csv), so a single workload is used across all four.
//
// Fixture provenance mirrors the main matrix's "a neutral real producer authors it, everyone reads
// the same file" rule: xlsx reuses the ExcelJS-authored fixture; xlsb/ods are authored by SheetJS
// (a real producer that can WRITE those binary/ODF containers); csv is authored directly (it is just
// text — no producer bias). The SheetJS-reads-a-SheetJS-file caveat is the same one ExcelJS→xlsx
// already carries, and python-calamine (independent) anchors the xlsb/ods numbers.

import { existsSync, writeFileSync } from "node:fs";
import { buildDataset } from "./workloads.mjs";

// The one workload used across formats. `numbers` is the cleanest: identical IEEE doubles in every
// container (a 2-decimal value round-trips through xlsx/xlsb/ods binary AND csv's decimal text), so
// every reader materializes the same values and no shared-string encoding difference clouds it.
export const FORMAT_WORKLOAD = "numbers";

// Per-format reader lanes: which libraries can read each container (ExcelJS/openpyxl can't do
// xlsb/ods; python-calamine can't do csv — those lanes just don't list them).
export const FORMAT_READERS = {
	xlsx: [
		{ id: "openjsxl", label: "openjsxl" },
		{ id: "exceljs", label: "ExcelJS" },
		{ id: "xlsx", label: "SheetJS" },
	],
	xlsb: [
		{ id: "openjsxl", label: "openjsxl" },
		{ id: "xlsx", label: "SheetJS" },
	],
	ods: [
		{ id: "openjsxl", label: "openjsxl" },
		{ id: "xlsx", label: "SheetJS" },
	],
	csv: [
		{ id: "openjsxl", label: "openjsxl" },
		{ id: "exceljs", label: "ExcelJS" },
		{ id: "xlsx", label: "SheetJS" },
	],
};

export const FORMATS = Object.keys(FORMAT_READERS);

/**
 * Author (once, cached) the cross-format read fixture for `format` at `size`, returning its path.
 * xlsx returns the ExcelJS-authored path the main matrix already cached; csv is written directly;
 * xlsb/ods are written by SheetJS. Regenerated only when absent.
 */
export async function ensureFormatFixture(cache, size, format, xlsxPath) {
	if (format === "xlsx") return xlsxPath;
	const path = `${cache}read-${FORMAT_WORKLOAD}-${size.key}.${format}`;
	if (existsSync(path)) return path;
	const dataset = buildDataset(FORMAT_WORKLOAD, size.rows);
	if (format === "csv") {
		writeFileSync(path, datasetToCsv(dataset));
		return path;
	}
	const XLSX = await import("xlsx");
	const ws = XLSX.utils.aoa_to_sheet(dataset);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "Bench");
	const out = XLSX.write(wb, { type: "buffer", bookType: format, compression: true });
	writeFileSync(path, new Uint8Array(out));
	return path;
}

// Serialize the `numbers` dataset as CSV. Plain numbers need no quoting (RFC 4180), so a bare join is
// correct and keeps the fixture a faithful, producer-neutral text file.
function datasetToCsv(dataset) {
	let out = "";
	for (const row of dataset) out += `${row.join(",")}\n`;
	return out;
}
