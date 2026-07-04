# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

Fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) reader **and writer** for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge. This is the package to install; it re-exports
the [`@openjsxl/core`](https://www.npmjs.com/package/@openjsxl/core) engine.

"Fast" is measured, not asserted — on a 1M-cell sheet, ~2–3× the read throughput of ExcelJS/SheetJS
at a fraction of the memory ([benchmarks](https://github.com/joaquimserafim/openjsxl/blob/main/docs/benchmarks.md)).

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

See the [project README](https://github.com/joaquimserafim/openjsxl#readme) for the full guide,
design notes, and roadmap.

## License

MIT
