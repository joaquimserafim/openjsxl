// 01 — Open a workbook and read individual, typed cells.
//
//   node 01-read-cells.mjs        (from ./examples)
//   pnpm --filter openjsxl-examples read
//
// Cells are a discriminated union: `string | number | boolean | date | error | empty`.
// Narrowing on `cell.type` gives a correctly typed `cell.value` (a Date for `date`, etc.).

import { readFile } from 'node:fs/promises'
import { openXlsx } from 'openjsxl'

const bytes = await readFile(new URL('./data/sample.xlsx', import.meta.url))
const wb = await openXlsx(bytes)

console.log('Sheets:', wb.sheets.map((s) => s.name).join(', '))

const sales = wb.sheet('Sales')

for (const ref of ['A1', 'B2', 'C2', 'D2', 'E2']) {
	const cell = sales.cell(ref)
	// `E2` is a formula — the reader returns its cached result, typed like any other value.
	console.log(`${ref}: ${cell.type.padEnd(7)} ${JSON.stringify(cell.value)}`)
}
