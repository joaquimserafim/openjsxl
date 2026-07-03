# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![CI](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml/badge.svg)](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **reader + writer — pre-1.0.** Read typed cells, styles, number formats, merged ranges,
> hyperlinks, and comments; stream large sheets in roughly constant memory; get typed errors on
> malformed input. Write `.xlsx` from plain data with `writeXlsx`, and read → modify → write with
> `workbookToInput`. **New in 0.4:** cell styles (fonts, fills, borders, alignment, number
> formats) read *and* write, sheet geometry (column widths, row heights, frozen panes), and
> structural metadata (merged ranges, hyperlinks, sheet visibility) — all of it round-trips.
> Published on npm. Built in the open, plan-first — see the [roadmap](./ROADMAP.md) and
> [implementation plan](./IMPLEMENTATION.md).

## Quick start

`.xlsx` → JSON, in well under 50 lines:

```ts
import { openXlsx } from 'openjsxl';
import { readFile } from 'node:fs/promises';

// Open a workbook from bytes (Uint8Array or ArrayBuffer).
const wb = await openXlsx(await readFile('data.xlsx'));

// Read one cell — it's typed, so `cell.type` narrows `cell.value`.
const a1 = wb.sheet('Sheet1').cell('A1');
console.log(a1.type, a1.value); // e.g. "string" "hello"

// Or turn a whole sheet into JSON records keyed by column letter.
const sheet = wb.sheet(wb.sheets[0].name);
const rows = [];
for await (const row of sheet.rows()) {
	const record = {};
	for (const cell of row.cells) {
		record[cell.ref.replace(/\d+$/, '')] = cell.value;
	}
	rows.push(record);
}
console.log(JSON.stringify(rows, null, 2));
```

Cells are a discriminated union — `string`, `number`, `boolean`, `date`, `error`, or
`empty` — so narrowing on `cell.type` gives you a correctly typed `cell.value`.

### Large files

`openXlsx` decompresses each sheet so `cell('A1')` is random-access and synchronous. For
sheets too big to hold in memory, `streamSheetRows` reads row-at-a-time with roughly constant
memory — the worksheet is never materialized whole:

```ts
import { streamSheetRows } from 'openjsxl';
import { readFile } from 'node:fs/promises';

const bytes = await readFile('huge.xlsx');
for await (const row of streamSheetRows(bytes /*, 'Sheet1' */)) {
	// one row at a time; previous rows are already freed
	process.stdout.write(`${row.index}: ${row.cells.length} cells\n`);
}
```

### Styles, metadata & typed errors

Beyond values, a worksheet exposes each cell's resolved style and the metadata Excel writers
attach — number formats, merged ranges, hyperlinks, comments, sheet geometry, and the declared
used range. Number formats resolve through the cell's own style, then the column (`<col>`) and
row defaults, so a date column reads as dates even when its cells carry no style of their own.
Malformed input throws a typed `XlsxError` with a discriminating `code`, never a bare
`TypeError` from a corrupt file.

```ts
import { openXlsx, XlsxError } from 'openjsxl';
import { readFile } from 'node:fs/promises';

let wb;
try {
	// `maxPartBytes` caps any single decompressed part — an opt-in zip-bomb guard.
	wb = await openXlsx(await readFile('report.xlsx'), { maxPartBytes: 50_000_000 });
} catch (err) {
	if (err instanceof XlsxError) {
		// code: 'not-a-zip' | 'not-xlsx' | 'missing-part' | 'corrupt-zip' | 'part-too-large' | …
		console.error(`could not read xlsx (${err.code}): ${err.message}`);
		process.exit(1);
	}
	throw err;
}

const sheet = wb.sheet(wb.sheets[0].name);

sheet.style('B2'); // { font?, fill?, border?, alignment?, numberFormat? } | undefined
sheet.numberFormat('C1'); // "mm-dd-yy" — the format code, independent of the value
sheet.dimension; // "A1:E2" | undefined (the declared used range)
sheet.mergedCells; // ["A1:B1", "A2:A4", …]
sheet.hyperlinks; // [{ ref: "A1", target: "https://…", tooltip?, location?, display? }, …]
sheet.comments; // [{ ref: "A1", author: "Ada", text: "note" }, …]
sheet.columns; // [{ min: 2, max: 3, width: 25.5, hidden? }, …] — column geometry
sheet.rowProperties; // Map<row, { height?, hidden? }>
sheet.freeze; // { rows?, cols? } | undefined — the frozen pane
sheet.state; // "visible" | "hidden" | "veryHidden" (sheet.visible is the boolean)
```

## Writing

Describe a workbook as plain data and `writeXlsx` returns the `.xlsx` bytes. The cell type is
inferred from each JavaScript value — `string`, `number`, `boolean`, `Date` (written as a
date-formatted serial), and `null`/`undefined` for an empty cell:

```ts
import { writeXlsx } from 'openjsxl';
import { writeFile } from 'node:fs/promises';

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
});

await writeFile('report.xlsx', bytes); // opens cleanly in Excel and LibreOffice
```

The output is deterministic (identical input → identical bytes), strings are written inline (no
shared-strings table), and input the format can't represent — no sheets, a bad or duplicate sheet
name, a non-finite number, an invalid `Date`, or a string with XML-illegal characters — throws a
typed `XlsxError` with `code: 'invalid-input'` rather than producing a file Excel must repair.

### Styles & layout (0.4)

A cell can also be `{ value, style }` — the style shape is exactly what `sheet.style(ref)`
returns, so styles pass straight through a round trip. Sheets take column widths, row heights,
frozen panes, merged ranges, hyperlinks, and a visibility state:

```ts
const bold = { font: { bold: true }, fill: { patternType: 'solid', fgColor: { rgb: 'FFDDEBF7' } } };

const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [
				[
					{ value: 'Item', style: bold },
					{ value: 'Total', style: bold },
				],
				['Apples', { value: 1234.5, style: { numberFormat: '#,##0.00' } }],
			],
			columns: [{ min: 1, max: 1, width: 18 }], // widen column A
			freeze: { rows: 1 }, // keep the header visible
			merges: ['A3:B3'],
			hyperlinks: [{ ref: 'A2', target: 'https://example.com/apples', tooltip: 'docs' }],
		},
		{ name: 'Internal', rows: [['scratch']], state: 'hidden' },
	],
});
```

### Read → modify → write

`workbookToInput` turns an open `Workbook` back into writer input, so you can round-trip a file:

```ts
import { openXlsx, workbookToInput, writeXlsx } from 'openjsxl';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await openXlsx(await readFile('in.xlsx'));
const input = await workbookToInput(wb);
input.sheets[0].rows.push(['appended', 'row']); // tweak the plain data
await writeFile('out.xlsx', await writeXlsx(input));
```

The round trip is **lossless for values, types, sheet names/order, styles, geometry, and
structural metadata**:

| Round-trips losslessly | Not carried (yet) |
| --- | --- |
| string, number, boolean, `Date` values | comments (planned for 0.5) |
| number formats — built-in & custom codes | formulas (only the cached value survives) |
| fonts, fills, borders, alignment | error cells (written as their text) |
| colors: rgb, indexed, theme + tint (raw) | |
| empty cells (sparse), incl. styled blanks | |
| column widths, row heights, hidden, freeze | |
| merged ranges & hyperlinks | |
| sheet visibility (hidden / veryHidden) | |
| sheet names & tab order | |

Three documented flattenings (values stay exact; internal spelling normalizes): row/column
*default* styles resolve into per-cell styles (each cell keeps its effective format); files
authored under a **custom theme** keep their `{theme, tint}` color indexes but re-render against
the standard Office theme after a rewrite (`rgb`/`indexed` colors are unaffected); and a
tolerated non-canonical cell ref spelling (e.g. lowercase `a1`) re-emits canonically (`A1`).

## Why

JavaScript has no Excel library that is, all at once, maintained, permissively licensed,
published on npm, ESM/TypeScript-first, dependency-free, and fast. SheetJS's public **npm**
release is frozen on a 2022 build that predates its own security fix — current releases ship
only from a vendor CDN, and styled write is a paid tier; ExcelJS is MIT and full-featured but
effectively unmaintained (no release since 2023, hundreds of open issues). openjsxl aims for
that empty square — taking the speed lessons of Python's
[`python-calamine`](https://pypi.org/project/python-calamine/) and growing toward the
capability of [`openpyxl`](https://pypi.org/project/openpyxl/).

## Performance

"Fast" is measured, not asserted. Against the JS incumbents on an Apple M2 Pro (Node 24), a
**1-million-cell** sheet (median wall-time · peak RSS, each library in an isolated process):

| 1M cells | openjsxl | ExcelJS `4.4.0` | SheetJS `0.18.5` |
| --- | --- | --- | --- |
| **read** | **0.71 s · 205 MB** | 1.6 s · 717 MB | 2.2 s · 547 MB |
| **write** | **0.73 s · 389 MB** | 3.2 s · 1.5 GB | 2.5 s · 566 MB |

Writing with `streamXlsx` from a lazy row source holds memory roughly **flat (~100 MB)** no matter
the row count. The full matrix (10k / 100k / 1M cells, numbers / strings / styled, read + write),
the methodology, and out-of-band **openpyxl** / **python-calamine** reference numbers are in
[`docs/benchmarks.md`](./docs/benchmarks.md) — reproduce it end-to-end with `pnpm bench`.
_(Measured 2026-07-03; every "fast" in these docs traces back to this table.)_

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
