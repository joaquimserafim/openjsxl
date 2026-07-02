// 02 — Turn a sheet into JSON records, using row 1 as the header.
//
//   node 02-sheet-to-json.mjs
//   pnpm --filter openjsxl-examples json
//
// Rows and cells are sparse: only populated cells appear, each carrying its own A1 `ref`.

import { readFile } from "node:fs/promises"
import { openXlsx } from "openjsxl"

const wb = await openXlsx(await readFile(new URL("./data/sample.xlsx", import.meta.url)))
const sheet = wb.sheet("Sales")

const rows = []
for await (const row of sheet.rows()) rows.push(row)

// Column letter → header label, from the first row.
const colOf = (ref) => ref.replace(/\d+$/, "")
const header = {}
for (const cell of rows[0].cells) header[colOf(cell.ref)] = cell.value

// Map every following row to an object keyed by header label.
const records = rows.slice(1).map((row) => {
	const record = {}
	for (const cell of row.cells) record[header[colOf(cell.ref)] ?? colOf(cell.ref)] = cell.value
	return record
})

console.log(JSON.stringify(records, null, 2))
