# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

Fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) reader **and writer** for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge. This is the package to install; it re-exports
the [`@openjsxl/core`](https://www.npmjs.com/package/@openjsxl/core) engine.

"Fast" is measured, not asserted — on a 1M-cell sheet, ~2–3× the read throughput of ExcelJS/SheetJS
at a fraction of the memory, and it installs in **~0.3 MB with zero third-party dependencies** (vs
ExcelJS's 34 MB / 96 packages and SheetJS's 14 MB / 8)
([benchmarks](https://github.com/joaquimserafim/openjsxl/blob/main/docs/benchmarks.md)).

```sh
npm install openjsxl
```

```ts
import { openXlsx } from 'openjsxl';
import { readFile } from 'node:fs/promises';

const wb = await openXlsx(await readFile('data.xlsx'));

// Typed cells: narrowing on `cell.type` gives a correctly typed `cell.value`.
const a1 = wb.sheet('Sheet1').cell('A1');
console.log(a1.type, a1.value); // e.g. "string" "hello"

// Stream a whole sheet, row at a time.
for await (const row of wb.sheet(wb.sheets[0].name).rows()) {
	console.log(
		row.index,
		row.cells.map((c) => c.value),
	);
}
```

For very large sheets use `streamSheetRows` (constant memory); a worksheet also exposes
`style(ref)`, `numberFormat`, `formula(ref)`, `dimension`, `mergedCells`, `hyperlinks`,
`comments`, `columns`, `rowProperties`, `freeze`, `state`/`visible`, and async `images()`
(anchored pictures — raw bytes, media type, and cell + EMU anchor); the workbook resolves
theme colors to ARGB with `resolveColor`; and the reader throws a typed `XlsxError` (with a
discriminating `code`) on malformed input.

**Writing:** describe a workbook as plain data and get back `.xlsx` bytes — cell types are
inferred from the JS values. Cells can carry styles (`{ value, style }` — the same shape
`style(ref)` returns) or formulas (`{ formula, value? }`), and sheets take column widths, row
heights, frozen panes, merged ranges, hyperlinks, comments (Excel-visible), a visibility
state, and `images` (the same records `images()` returns — png, jpeg, gif, bmp, tiff, webp,
emf, wmf). For exports too big to buffer, `streamXlsx` streams the same input shape from lazy
(sync or async) row sources with roughly constant memory — constant in *rows*; embedded image
bytes are held until the media parts flush at stream end:

```ts
import { writeXlsx } from 'openjsxl';

const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [
				[{ value: 'Item', style: { font: { bold: true } } }, 'Added'],
				['Apples', new Date('2024-01-15')],
			],
			freeze: { rows: 1 },
			hyperlinks: [{ ref: 'A2', target: 'https://example.com/apples' }],
		},
	],
});
```

`workbookToInput` turns an open `Workbook` back into writer input for read → modify → write —
values, types, styles, formulas, comments, pictures, custom themes, geometry, merges,
hyperlinks, and sheet visibility all round-trip.

**Reading other formats (0.7):** openjsxl writes `.xlsx` but reads more — `openXlsb` (Excel
Binary Workbook), `openOds` (OpenDocument), and `openCsv` (delimited text) all return the SAME
`Workbook`, and `detectSpreadsheetFormat(bytes)` routes by content (`'xlsx' | 'xlsb' | 'ods' |
'csv' | undefined`; `.xlsm`/`.xltx` read as `'xlsx'`). Accessors a format can't express degrade,
never throw — `.xlsb`/`.ods` carry values, dates, merges (ods) and hyperlinks; `.csv` infers
numbers & booleans only (never dates). So "a user uploaded a spreadsheet" is one code path, and any
reader converts to `.xlsx` through the bridge.

**Formulas (0.8):** the reader keeps a formula's text and cached value; the opt-in
`openjsxl/formula` entry adds a zero-dependency engine that recomputes them —
`evaluateWorkbook(wb)` / `evaluateCell(wb, sheet, ref)` over 90+ built-in functions, plus your own
via `options.functions`. It's a separate import (the core bundle is unchanged whether or not you
use it); evaluation is read-only (it can supersede a stale cache), circular references resolve to a
`#CYCLE!` value instead of hanging, and volatile functions (`TODAY`/`RAND`) require an injected
clock/RNG so results stay deterministic.

See the [project README](https://github.com/joaquimserafim/openjsxl#readme) for the full guide,
design notes, and roadmap.

## License

MIT
