// 07 — Styles, geometry, and structural metadata (0.4): write a styled workbook, read it back.
//
//   node 07-styles-and-layout.mjs        (from ./examples)
//   pnpm --filter openjsxl-examples styles
//
// A cell can be `{ value, style }` — the style is EXACTLY what `sheet.style(ref)` returns, so
// styles pass through read → modify → write untouched. Sheets take column widths, row heights,
// frozen panes, merged ranges, hyperlinks, and a visibility state; all of it round-trips.

import { writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { openXlsx, writeXlsx } from "openjsxl"

const header = {
	font: { bold: true, color: { rgb: "FFFFFFFF" } },
	fill: { patternType: "solid", fgColor: { rgb: "FF4472C4" } },
	alignment: { horizontal: "center" },
}

const bytes = await writeXlsx({
	sheets: [
		{
			name: "Report",
			rows: [
				[
					{ value: "Item", style: header },
					{ value: "Total", style: header },
				],
				["Apples", { value: 1234.5, style: { numberFormat: "#,##0.00" } }],
				["Pears", { value: 0.176, style: { numberFormat: "0.0%" } }],
				["All fruit ->"], // A4:B4 is merged below; only the top-left cell holds a value
			],
			columns: [{ min: 1, max: 1, width: 18 }], // widen column A
			rowProperties: { 1: { height: 24 } }, // taller header row
			freeze: { rows: 1 }, // header stays visible while scrolling
			merges: ["A4:B4"],
			hyperlinks: [
				{ ref: "A2", target: "https://example.com/apples", tooltip: "product page" },
			],
		},
		{ name: "Internal", rows: [["scratch space"]], state: "hidden" },
	],
})

const out = join(tmpdir(), "openjsxl-styled.xlsx")
await writeFile(out, bytes)
console.log(`wrote ${bytes.length} bytes -> ${out}`)

// Read it back — every accessor returns exactly what was written.
const wb = await openXlsx(bytes)
const sheet = wb.sheet("Report")
console.log("A1 style     :", JSON.stringify(sheet.style("A1")))
console.log("B2 format    :", sheet.numberFormat("B2"), "->", sheet.cell("B2").value)
console.log("columns      :", JSON.stringify(sheet.columns))
console.log("freeze       :", JSON.stringify(sheet.freeze))
console.log("merges       :", JSON.stringify(sheet.mergedCells))
console.log("hyperlinks   :", JSON.stringify(sheet.hyperlinks))
console.log("sheet states :", wb.sheets.map((s) => `${s.name}=${s.state}`).join(", "))
