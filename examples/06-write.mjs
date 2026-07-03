// 06 — Write an .xlsx from plain data, and round-trip an existing one (read → modify → write).
//
//   node 06-write.mjs        (from ./examples)
//   pnpm --filter openjsxl-examples write
//
// `writeXlsx` infers the OOXML cell type from each JS value: string, number, boolean, Date (a
// date-formatted serial), and null/undefined for an empty cell. `workbookToInput` turns a workbook
// you just read back into that same plain-data shape, so you can read, tweak, and write it out.

import { readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { openXlsx, workbookToInput, writeXlsx } from "openjsxl";

// (1) Author a workbook from scratch.
const bytes = await writeXlsx({
	sheets: [
		{
			name: "Report",
			rows: [
				["Item", "Qty", "Price", "In stock", "Added"],
				["Apples", 120, 0.5, true, new Date("2024-01-15")],
				["Pears", 80, 0.75, false, new Date("2024-02-01")],
			],
		},
	],
});
const out = join(tmpdir(), "openjsxl-report.xlsx");
await writeFile(out, bytes);
console.log(`wrote ${bytes.length} bytes → ${out}`);

// Re-open what we just wrote — the reader types it back exactly.
const report = (await openXlsx(bytes)).sheet("Report");
for (const ref of ["A2", "B2", "D2", "E2"]) {
	const cell = report.cell(ref);
	console.log(`  ${ref}: ${cell.type.padEnd(7)} ${JSON.stringify(cell.value)}`);
}

// (2) Round-trip: read an existing workbook, append a row, write it back.
const sample = await openXlsx(await readFile(new URL("./data/sample.xlsx", import.meta.url)));
const input = await workbookToInput(sample);
const sales = input.sheets.find((s) => s.name === "Sales");
const rowNumber = sales.rows.length + 1; // 1-based row the appended data will occupy
sales.rows.push(["Totals", 200]); // plain data — the same shape writeXlsx accepts
const modified = await writeXlsx(input);

const appended = (await openXlsx(modified)).sheet("Sales").cell(`A${rowNumber}`);
console.log(
	`round-tripped Sales (+1 row) → ${modified.length} bytes; A${rowNumber} = ${JSON.stringify(appended.value)}`,
);
