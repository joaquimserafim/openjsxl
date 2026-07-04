# @openjsxl/core

[![npm version](https://img.shields.io/npm/v/@openjsxl/core?color=cb3837&logo=npm)](https://www.npmjs.com/package/@openjsxl/core)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen)](https://www.npmjs.com/package/@openjsxl/core?activeTab=dependencies)
[![license: MIT](https://img.shields.io/npm/l/@openjsxl/core?color=blue)](./LICENSE)

The zero-dependency OOXML engine behind [`openjsxl`](https://www.npmjs.com/package/openjsxl):
the `zip → xml → ooxml → reader` layers that turn an `.xlsx` into typed cells — and the mirror
`writer` layer that turns plain data back into `.xlsx` bytes. Built only on platform Web APIs
(`DecompressionStream`, `CompressionStream`, `TextDecoder`, …). No runtime dependencies.

**Most users should install [`openjsxl`](https://www.npmjs.com/package/openjsxl) instead** — it
re-exports everything here and is the stable public surface. Install `@openjsxl/core` directly
only if you want the engine without the facade.

```sh
npm install @openjsxl/core
```

## Usage

```ts
import { openXlsx, streamSheetRows, XlsxError } from '@openjsxl/core';
import { readFile } from 'node:fs/promises';

const wb = await openXlsx(await readFile('data.xlsx'));
const sheet = wb.sheet('Sheet1');

sheet.cell('A1'); // { ref, type, value } — narrow on `type` for a typed value
sheet.style('B2'); // { font?, fill?, border?, alignment?, numberFormat? } | undefined
sheet.numberFormat('C1'); // "mm-dd-yy" | undefined
sheet.formula('E1'); // "B1*2" | undefined — formula text (shared formulas come back translated)
sheet.mergedCells; // ["A1:B1", …]
sheet.freeze; // { rows?, cols? } | undefined — plus columns, rowProperties, comments, state, …
await sheet.images(); // [{ anchor, bytes, mime, name? }, …] — anchored pictures (lazy; media
// bytes decompress on first call, once per drawing part)
wb.resolveColor({ theme: 4, tint: 0.4 }); // "FF96B4D8" — theme color → ARGB, per this file's theme

// Constant-memory streaming for large sheets — one row at a time.
for await (const row of streamSheetRows(await readFile('huge.xlsx'))) {
	console.log(row.index, row.cells.length);
}
```

Malformed input throws a typed `XlsxError` with a discriminating `.code`
(`'not-a-zip' | 'not-xlsx' | 'missing-part' | 'corrupt-zip' | 'part-too-large' | …`), never a
bare `TypeError` from a corrupt file.

## Writing

```ts
import { writeXlsx, workbookToInput } from '@openjsxl/core';

// Author from plain data — cell types inferred from the JS values.
const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [
				['Item', 'Added'],
				['Apples', new Date('2024-01-15')],
			],
		},
	],
});

// Or read → modify → write.
const input = await workbookToInput(await openXlsx(bytes));
input.sheets[0].rows.push(['Pears', new Date('2024-02-01')]);
const updated = await writeXlsx(input);
```

Cells can carry styles (`{ value, style }` — the same shape `style(ref)` returns) or formulas
(`{ formula, value? }` — the cached value is what non-recalculating readers see), and sheets
take `columns` (widths), `rowProperties` (heights), `freeze`, `merges`, `hyperlinks`,
`comments` (written with the legacy VML part Excel needs to display them), a visibility
`state`, and `images` (the same `{ anchor, bytes, mime, name? }` records `images()` returns —
png, jpeg, gif, bmp, tiff, webp, emf, wmf; identical bytes dedupe into one media part). For
huge exports, `streamXlsx` accepts the same input shape with each sheet's `rows` as any
sync/async iterable and returns a `ReadableStream` — roughly constant memory at any row count
(constant in *rows*; embedded image bytes are held, by reference, until the media parts flush
at stream end). Output is deterministic; unrepresentable input (no sheets, bad/duplicate sheet
name, non-finite number, invalid `Date`, XML-illegal characters, malformed or overlapping
merges) throws `XlsxError` with `code: 'invalid-input'`. The round trip is lossless for values,
types, sheet names/order, styles, formulas, comments, pictures, custom themes, geometry,
merges, hyperlinks, and visibility.

## Exports

- **Reader:** `openXlsx`, `streamSheetRows`, `Workbook`, `Worksheet`, `ReadOptions`
- **Writer:** `writeXlsx`, `streamXlsx`, `workbookToInput`, `WorkbookInput`, `SheetInput`,
  `CellInput`, `StyledCell`, `CellValue`, `WriteOptions`, `StreamWorkbookInput`,
  `StreamSheetInput`, `StreamRows`
- **Errors:** `XlsxError`, `XlsxErrorCode`
- **Types:** `Row`, `Cell`, `CellType`, `CellStyle` (+ font/fill/border/alignment/color parts),
  `ColumnProps`, `RowProps`, `FreezePane`, `Comment`, `Hyperlink`, `SheetInfo`, `SheetState`,
  `CellRef`, `SheetImage`, `ImageAnchor`, `AnchorPoint`
- **A1 & dates:** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef`, `serialToDate`,
  `dateToSerial`

Full guide, design notes, and roadmap: <https://github.com/joaquimserafim/openjsxl>

## License

MIT
