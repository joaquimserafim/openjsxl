# openjsxl

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **v0.1 — reader MVP.** Open an `.xlsx` and read typed cells today; not yet
> published to npm. Built in the open, plan-first — see the [roadmap](./ROADMAP.md) and
> [implementation plan](./IMPLEMENTATION.md).

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

Cells are a discriminated union — `string`, `number`, `boolean`, `error`, or `empty` (and
`date` once F2.1 lands) — so narrowing on `cell.type` gives you a correctly typed `cell.value`.

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
