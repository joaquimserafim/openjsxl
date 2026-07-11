// 11 — Read other formats (0.7): detect and read .xlsb, .ods, and .csv into the SAME Workbook API.
//
//   node 11-other-formats.mjs
//   pnpm --filter openjsxl-examples formats
//
// openjsxl WRITES .xlsx, but it READS more: `openXlsb` (Excel Binary Workbook), `openOds`
// (OpenDocument / LibreOffice), and `openCsv` (delimited text) all return the SAME `Workbook` as
// `openXlsx` — typed cells, the same accessors, the same discriminated-union `Cell`.
// `detectSpreadsheetFormat` sniffs the bytes, so "a user uploaded a spreadsheet" is ONE code path
// regardless of dialect, and any reader is a converter to .xlsx through the bridge (`workbookToInput`
// → `writeXlsx`). Accessors a format can't express degrade to []/undefined.

import { readFile } from "node:fs/promises";
import {
	detectSpreadsheetFormat,
	openCsv,
	openOds,
	openXlsb,
	openXlsx,
	workbookToInput,
	writeXlsx,
} from "openjsxl";

// Detect the format from the bytes, then open with the matching reader — the caller never hard-codes
// "this is an xlsb". `openCsv` is synchronous (no container to decompress); the others are async.
async function openAnything(bytes) {
	const format = await detectSpreadsheetFormat(bytes);
	switch (format) {
		case "xlsx":
			return { format, wb: await openXlsx(bytes) };
		case "xlsb":
			return { format, wb: await openXlsb(bytes) };
		case "ods":
			return { format, wb: await openOds(bytes) };
		case "csv":
			return { format, wb: openCsv(bytes) };
		default:
			throw new Error("unrecognized spreadsheet format");
	}
}

const show = (file, format, wb) => {
	const sheet = wb.sheet(wb.sheets[0].name);
	console.log(
		`\n${file} — detected ${format}; sheets: ${wb.sheets.map((s) => s.name).join(", ")}`,
	);
	for (const ref of ["A1", "B1", "C1"]) {
		const cell = sheet.cell(ref);
		console.log(`  ${ref}: ${cell.type} = ${JSON.stringify(cell.value)}`);
	}
};

// One loop over three dialects — the reader is chosen by the bytes, not by the file extension.
for (const file of ["other-formats.xlsb", "other-formats.ods", "other-formats.csv"]) {
	const bytes = await readFile(new URL(`./data/${file}`, import.meta.url));
	const { format, wb } = await openAnything(bytes);
	show(file, format, wb);
}

// Any reader is a converter: turn the .xlsb into .xlsx bytes through the bridge, then read it back.
const xlsbBytes = await readFile(new URL("./data/other-formats.xlsb", import.meta.url));
const { wb: xlsb } = await openAnything(xlsbBytes);
const asXlsx = await writeXlsx(await workbookToInput(xlsb));
const roundTripped = await openXlsx(asXlsx);
console.log(
	`\nconverted xlsb → xlsx (${asXlsx.length} bytes); re-read A1 =`,
	JSON.stringify(roundTripped.sheet(roundTripped.sheets[0].name).cell("A1").value),
);
