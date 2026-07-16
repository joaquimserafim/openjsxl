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
sheet.tables; // [{ name, ref, columns, headerRow, totalsRow, style? }, …] — defined tables (0.9)
sheet.dataValidations; // [{ sqref, type, operator?, formula1?, … }, …] — dropdowns/input rules (0.9)
sheet.conditionalFormatting; // [{ sqref, rules }, …] — highlight/colorScale/dataBar/iconSet (0.9)
wb.resolveColor({ theme: 4, tint: 0.4 }); // "FF96B4D8" — theme color → ARGB, per this file's theme

// Constant-memory streaming for large sheets — one row at a time.
for await (const row of streamSheetRows(await readFile('huge.xlsx'))) {
	console.log(row.index, row.cells.length);
}
```

Malformed input throws a typed `XlsxError` with a discriminating `.code`
(`'not-a-zip' | 'not-xlsx' | 'missing-part' | 'corrupt-zip' | 'part-too-large' | …`), never a
bare `TypeError` from a corrupt file. Reads are guarded against decompression bombs **by default** —
a 2 GiB per-part output ceiling plus a 300× compression-ratio cap (over an 8 MiB floor); raise or
disable either via `ReadOptions.maxPartBytes` / `maxCompressionRatio` (`Number.POSITIVE_INFINITY`
to turn one off).

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

// Or read → modify → write (the input is deeply readonly — spread, don't mutate).
const input = await workbookToInput(await openXlsx(bytes));
const sheets = input.sheets.map((sheet, i) =>
	i === 0 ? { ...sheet, rows: [...sheet.rows, ['Pears', new Date('2024-02-01')]] } : sheet,
);
const updated = await writeXlsx({ ...input, sheets });
```

Cells can carry styles (`{ value, style }` — the same shape `style(ref)` returns) or formulas
(`{ formula, value? }` — the cached value is what non-recalculating readers see), and sheets
take `columns` (widths), `rowProperties` (heights), `freeze`, `merges`, `hyperlinks`,
`comments` (written with the legacy VML part Excel needs to display them), a visibility
`state`, `images` (the same `{ anchor, bytes, mime, name? }` records `images()` returns —
png, jpeg, gif, bmp, tiff, webp, emf, wmf; identical bytes dedupe into one media part), and
`tables` / `dataValidations` / `conditionalFormatting` (0.9). For
huge exports, `streamXlsx` accepts the same input shape with each sheet's `rows` as any
sync/async iterable and returns a `ReadableStream` — roughly constant memory at any row count
(constant in *rows*; embedded image bytes are held, by reference, until the media parts flush
at stream end). `writeXlsx(input, { date1904: true })` selects the legacy 1904 date epoch.
Output is deterministic; unrepresentable input (no sheets, bad/duplicate sheet name,
non-finite number, invalid `Date`, malformed or overlapping merges) throws `XlsxError` with
`code: 'invalid-input'` — string *content* never throws: XML-illegal characters (controls,
lone surrogates) store via the `_xHHHH_` escape Excel itself uses, and round-trip. The round
trip is lossless for values, types, sheet names/order, styles, formulas, comments, pictures,
custom themes, geometry, merges, hyperlinks, visibility, defined names, tables, data
validations, conditional formatting, autofilter ranges, protection (sheet/workbook locks +
per-cell locked/hidden, password hashes verbatim), and print setup (margins, page setup,
print options, header/footer). Documented drops (never silent): in-cell rich text flattens to
plain text; an autofilter's per-column criteria/sort, row/column outline grouping, sheet tab
colors, document properties, pivot tables, external-workbook links, gradient fills, and threaded
-comment threading are not carried; and **VBA macros** — an `.xlsm` reads but rewrites to a plain
`.xlsx` without them (`Workbook.macroEnabled` flags a macro-enabled source). Full list: the root
README's fidelity table.

## Other formats (read-only)

openjsxl writes `.xlsx`, but reads more: `openXlsb` (Excel Binary Workbook), `openOds`
(OpenDocument), and `openCsv` (delimited text) return the SAME `Workbook` as `openXlsx`, and
`detectSpreadsheetFormat(bytes)` → `'xlsx' | 'xlsb' | 'ods' | 'csv' | undefined` routes by content
(container formats by their package; CSV by a documented text heuristic). Accessors a format can't
express degrade (`style()` → `undefined`, `mergedCells` → `[]`), never throw. `.xlsb`/`.ods` carry
values, dates, merges (ods) and hyperlinks; `.csv` infers numbers & booleans only (never dates).
Any of them converts to `.xlsx` through the bridge (`workbookToInput` → `writeXlsx`).

## Exports

- **Reader:** `openXlsx`, `streamSheetRows`, `Workbook`, `Worksheet`, `ReadOptions`
- **Other-format readers:** `openXlsb`, `openOds`, `openCsv` (+ `CsvReadOptions`),
  `detectSpreadsheetFormat` (+ `SpreadsheetFormat`)
- **Writer:** `writeXlsx`, `streamXlsx`, `workbookToInput`, `WorkbookInput`, `SheetInput`,
  `CellInput`, `StyledCell`, `CellValue`, `WriteOptions`, `StreamWorkbookInput`,
  `StreamSheetInput`, `StreamRows`
- **Errors:** `XlsxError`, `XlsxErrorCode`
- **Types:** `Row`, `Cell`, `CellType`, `CellStyle` (+ font/fill/border/alignment/color parts),
  `ColumnProps`, `RowProps`, `FreezePane`, `Comment`, `Hyperlink`, `SheetInfo`, `SheetState`,
  `CellRef`, `SheetImage`, `ImageAnchor`, `AnchorPoint`, `TableInfo`, `TableColumn`,
  `TableStyleInfo`, `DataValidation` (+ `DataValidationType`/`DataValidationOperator`/
  `DataValidationErrorStyle`), `ConditionalFormatting`, `ConditionalFormattingRule`, `DxfStyle`,
  `DxfFill`, `Cfvo`
- **A1 & dates:** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef`, `serialToDate`,
  `dateToSerial`
- **Formulas (opt-in, `@openjsxl/core/formula`):** `parseFormula`, `evaluateWorkbook`,
  `evaluateCell`, `FormulaError` (+ `FormulaErrorCode`), `EvaluateOptions`, `FunctionSpec`
  (+ `EagerFunctionSpec`/`LazyFunctionSpec`/`EvalContext`/`ArgThunk`), `EvalValue`/`ScalarValue`,
  `FormulaErrorValue`/`ErrorCode`/`errorValue`/`isErrorValue`, `RangeView`/`isRangeView`, and the
  AST node types. A separate entry point — importing it never changes the core `"."` bundle. ~90
  built-in functions; register your own through `options.functions`.

Full guide, design notes, and roadmap: <https://github.com/joaquimserafim/openjsxl>

## License

MIT
