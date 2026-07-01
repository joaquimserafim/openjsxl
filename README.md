# openjsxl

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **reader — hardened, pre-1.0.** Read typed cells, number formats, merged ranges,
> hyperlinks, and comments; stream large sheets in roughly constant memory; get typed errors on
> malformed input. A writer is next. First npm release (`v0.1.0`) is imminent. Built in the open,
> plan-first — see the [roadmap](./ROADMAP.md) and [implementation plan](./IMPLEMENTATION.md).

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

## Why

JavaScript has no Excel library that is, all at once, maintained, permissively licensed,
published on npm, ESM/TypeScript-first, dependency-free, and fast. SheetJS is frozen on a
vulnerable npm build behind a CDN and a paid tier; ExcelJS is effectively unmaintained.
openjsxl aims for that empty square — taking the speed lessons of Python's
[`python-calamine`](https://pypi.org/project/python-calamine/) and growing toward the
capability of [`openpyxl`](https://pypi.org/project/openpyxl/).

## Approach

- **Read first, write later.** A fast, correct reader earns trust before we ship a writer.
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
