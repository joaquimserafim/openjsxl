// 04 — Metadata a worksheet exposes beyond cell values.
//
//   node 04-metadata.mjs
//   pnpm --filter openjsxl-examples metadata
//
// Number formats resolve through the cell's own style, then the column (`<col>`) and row
// defaults — so column C reads as a date format even though its cells set no style of their own.

import { readFile } from "node:fs/promises"
import { openXlsx } from "openjsxl"

const wb = await openXlsx(await readFile(new URL("./data/sample.xlsx", import.meta.url)))
const sheet = wb.sheet("Sales")

console.log("dimension      :", sheet.dimension)
console.log("C2 numberFormat:", sheet.numberFormat("C2"), "(inherited from the column)")
console.log("mergedCells    :", sheet.mergedCells)
console.log("hyperlinks     :", sheet.hyperlinks)
console.log("comments       :", sheet.comments)
console.log("sheet visibility:")
for (const info of wb.sheets) {
	console.log(`  ${info.name}: ${info.visible ? "visible" : "hidden"}`)
}
