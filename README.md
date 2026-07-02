# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![CI](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml/badge.svg)](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **reader + writer — pre-1.0.** Read typed cells, number formats, merged ranges,
> hyperlinks, and comments; stream large sheets in roughly constant memory; get typed errors on
> malformed input. **New in 0.3:** write `.xlsx` from plain data with `writeXlsx`, and read →
> modify → write with `workbookToInput`. Published on npm. Built in the open, plan-first — see the
> [roadmap](./ROADMAP.md) and [implementation plan](./IMPLEMENTATION.md).

## Quick start

`.xlsx` → JSON, in well under 50 lines:

```ts
import { openXlsx } from 'openjsxl'
import { readFile } from 'node:fs/promises'

// Open a workbook from bytes (Uint8Array or ArrayBuffer).
const wb = await openXlsx(await readFile('data.xlsx'))

// Read one cell — it's typed, so `cell.type` narrows `cell.value`.
const a1 = wb.sheet('Sheet1').cell('A1')
console.log(a1.type, a1.value) // e.g. "string" "hello"

// Or turn a whole sheet into JSON records keyed by column letter.
const sheet = wb.sheet(wb.sheets[0].name)
const rows = []
for await (const row of sheet.rows()) {
	const record = {}
	for (const cell of row.cells) {
		record[cell.ref.replace(/\d+$/, '')] = cell.value
	}
	rows.push(record)
}
console.log(JSON.stringify(rows, null, 2))
```

Cells are a discriminated union — `string`, `number`, `boolean`, `date`, `error`, or
`empty` — so narrowing on `cell.type` gives you a correctly typed `cell.value`.

### Large files

`openXlsx` decompresses each sheet so `cell('A1')` is random-access and synchronous. For
sheets too big to hold in memory, `streamSheetRows` reads row-at-a-time with roughly constant
memory — the worksheet is never materialized whole:

```ts
import { streamSheetRows } from 'openjsxl'
import { readFile } from 'node:fs/promises'

const bytes = await readFile('huge.xlsx')
for await (const row of streamSheetRows(bytes /*, 'Sheet1' */)) {
	// one row at a time; previous rows are already freed
	process.stdout.write(`${row.index}: ${row.cells.length} cells\n`)
}
```

### Metadata, number formats & typed errors

Beyond values, a worksheet exposes the metadata Excel writers attach — number formats, merged
ranges, hyperlinks, comments, and the declared used range. Number formats resolve through the
cell's own style, then the column (`<col>`) and row defaults, so a date column reads as dates
even when its cells carry no style of their own. Malformed input throws a typed `XlsxError` with
a discriminating `code`, never a bare `TypeError` from a corrupt file.

```ts
import { openXlsx, XlsxError } from 'openjsxl'
import { readFile } from 'node:fs/promises'

let wb
try {
	// `maxPartBytes` caps any single decompressed part — an opt-in zip-bomb guard.
	wb = await openXlsx(await readFile('report.xlsx'), { maxPartBytes: 50_000_000 })
} catch (err) {
	if (err instanceof XlsxError) {
		// code: 'not-a-zip' | 'not-xlsx' | 'missing-part' | 'corrupt-zip' | 'part-too-large' | …
		console.error(`could not read xlsx (${err.code}): ${err.message}`)
		process.exit(1)
	}
	throw err
}

const sheet = wb.sheet(wb.sheets[0].name)

sheet.numberFormat('C1') // "mm-dd-yy" — the format code, independent of the value
sheet.dimension          // "A1:E2" | undefined (the declared used range)
sheet.mergedCells        // ["A1:B1", "A2:A4", …]
sheet.hyperlinks         // [{ ref: "A1", target: "https://…", tooltip?, location?, display? }, …]
sheet.comments           // [{ ref: "A1", author: "Ada", text: "note" }, …]
sheet.visible            // false for hidden / very-hidden sheets
```

## Writing

Describe a workbook as plain data and `writeXlsx` returns the `.xlsx` bytes. The cell type is
inferred from each JavaScript value — `string`, `number`, `boolean`, `Date` (written as a
date-formatted serial), and `null`/`undefined` for an empty cell:

```ts
import { writeXlsx } from 'openjsxl'
import { writeFile } from 'node:fs/promises'

const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [
				['Item', 'Qty', 'Price', 'Added'],
				['Apples', 120, 0.5, new Date('2024-01-15')],
				['Pears', 80, 0.75, null],
			],
		},
	],
})

await writeFile('report.xlsx', bytes) // opens cleanly in Excel and LibreOffice
```

The output is deterministic (identical input → identical bytes), strings are written inline (no
shared-strings table), and input the format can't represent — no sheets, a bad or duplicate sheet
name, a non-finite number, an invalid `Date`, or a string with XML-illegal characters — throws a
typed `XlsxError` with `code: 'invalid-input'` rather than producing a file Excel must repair.

### Read → modify → write

`workbookToInput` turns an open `Workbook` back into writer input, so you can round-trip a file:

```ts
import { openXlsx, workbookToInput, writeXlsx } from 'openjsxl'
import { readFile, writeFile } from 'node:fs/promises'

const wb = await openXlsx(await readFile('in.xlsx'))
const input = await workbookToInput(wb)
input.sheets[0].rows.push(['appended', 'row']) // tweak the plain data
await writeFile('out.xlsx', await writeXlsx(input))
```

The round trip is **lossless for values, types, sheet names/order, and cell styles**:

| Round-trips losslessly | Not carried (yet) |
| --- | --- |
| string, number, boolean, `Date` values | merged ranges, hyperlinks, comments |
| number formats — built-in & custom codes | formulas (only the cached value survives) |
| fonts, fills, borders, alignment | error cells (written as their text) |
| colors: rgb, indexed, theme + tint (raw) | sheet visibility |
| empty cells (sparse), incl. styled blanks | |
| sheet names & tab order | |

Two documented flattenings: row/column *default* styles resolve into per-cell styles (each cell
keeps its effective format), and files authored under a **custom theme** keep their `{theme, tint}`
color indexes but re-render against the standard Office theme after a rewrite — `rgb`/`indexed`
colors are unaffected.

## Why

JavaScript has no Excel library that is, all at once, maintained, permissively licensed,
published on npm, ESM/TypeScript-first, dependency-free, and fast. SheetJS is frozen on a
vulnerable npm build behind a CDN and a paid tier; ExcelJS is effectively unmaintained.
openjsxl aims for that empty square — taking the speed lessons of Python's
[`python-calamine`](https://pypi.org/project/python-calamine/) and growing toward the
capability of [`openpyxl`](https://pypi.org/project/openpyxl/).

## Approach

- **Read first, then write.** A fast, correct reader earned trust first; the writer (0.3) is its
  mirror image — every byte it emits reads back through the reader.
- **Zero runtime dependencies.** Zip inflate/deflate comes from the platform
  (`DecompressionStream` / `CompressionStream`), strings from `TextEncoder`/`TextDecoder`.
- **Layered & swappable.** `zip → xml → ooxml → reader`, with the hot path behind an
  interface so a native (napi-rs / WASM `calamine`) backend can slot in later.

## Packages

| Package | Description |
| --- | --- |
| [`openjsxl`](./packages/openjsxl) | Public facade — the package users install |
| [`@openjsxl/core`](./packages/core) | The zero-dependency OOXML engine |
| `@openjsxl/fixtures` | Private test corpus + generator |

## Development

Requires Node ≥ 24 and pnpm.

```sh
pnpm install
pnpm check       # Biome lint + format check
pnpm typecheck   # tsc, no emit
pnpm test        # Vitest
pnpm fixtures    # regenerate programmatic test fixtures
```

## License

MIT
