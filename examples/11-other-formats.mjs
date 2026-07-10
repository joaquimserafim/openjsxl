// 11 — Read other formats (0.7): .xlsb and .ods into the SAME Workbook API.
//
//   node 11-other-formats.mjs
//   pnpm --filter openjsxl-examples formats
//
// openjsxl writes .xlsx, but it reads more: `openXlsb` (Excel Binary Workbook) and `openOds`
// (OpenDocument / LibreOffice) both return the SAME `Workbook` as `openXlsx` — typed cells, the
// same accessors, the same discriminated-union `Cell`. So "a user uploaded a spreadsheet" is one
// code path regardless of dialect, and any reader becomes a converter to .xlsx through the bridge
// (`workbookToInput` → `writeXlsx`). Accessors a format can't express degrade to []/undefined.

import { readFile } from "node:fs/promises";
import { openOds, openXlsb, openXlsx, workbookToInput, writeXlsx } from "openjsxl";

const show = (label, wb) => {
	console.log(`\n${label} — sheets: ${wb.sheets.map((s) => s.name).join(", ")}`);
	const sheet = wb.sheet(wb.sheets[0].name);
	for (const ref of ["A1", "B1", "E1"]) {
		const cell = sheet.cell(ref);
		console.log(`  ${ref}: ${cell.type} = ${JSON.stringify(cell.value)}`);
	}
};

// Read an Excel Binary Workbook (.xlsb) — BIFF12 binary parts, same OPC container as .xlsx.
const xlsb = await openXlsb(await readFile(new URL("./data/other-formats.xlsb", import.meta.url)));
show("xlsb", xlsb);

// Read an OpenDocument spreadsheet (.ods) — a zip of XML, explicit cell types.
const ods = await openOds(await readFile(new URL("./data/other-formats.ods", import.meta.url)));
show("ods", ods);

// Any reader is a converter: turn the .xlsb into .xlsx bytes through the bridge, then read it back.
const asXlsx = await writeXlsx(await workbookToInput(xlsb));
const roundTripped = await openXlsx(asXlsx);
console.log(
	`\nconverted xlsb → xlsx (${asXlsx.length} bytes); re-read A1 =`,
	roundTripped.sheet(roundTripped.sheets[0].name).cell("A1").value,
);
