// 03 — Stream a sheet row-at-a-time with roughly constant memory.
//
//   node 03-stream-rows.mjs
//   pnpm --filter openjsxl-examples stream
//
// `streamSheetRows` never materializes the whole worksheet: it decompresses and tokenizes chunk
// by chunk, yielding each row then discarding it. Use it for sheets too large to hold in memory;
// use `openXlsx` when you need random `cell()` access. The second arg selects the sheet (default:
// the first). Here the sample is tiny, but the shape is identical for a million-row file.

import { readFile } from "node:fs/promises"
import { streamSheetRows } from "openjsxl"

const bytes = await readFile(new URL("./data/sample.xlsx", import.meta.url))

for await (const row of streamSheetRows(bytes, "Sales")) {
	console.log(
		`row ${row.index}:`,
		row.cells.map((c) => c.value),
	)
}
