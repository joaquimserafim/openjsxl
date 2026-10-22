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

**ESM-only** (no CommonJS build) and, on Node, **Node ≥ 24** — the platform Web APIs above ship
unflagged from 24 on. `import` is the entry point; `require('@openjsxl/core')` also resolves on
Node ≥ 24 via `require(esm)`.

**Security:** see the [security policy](https://github.com/joaquimserafim/openjsxl/blob/main/SECURITY.md)
for private vulnerability reporting, supported releases and limits when processing untrusted files.
It distinguishes published 1.0.0 behavior from protections on the unreleased 1.0.1 branch.

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
bare `TypeError` from a corrupt file. This is a uniform contract: **every public function reports
failure by throwing `XlsxError`** — the readers, the writer (`'invalid-input'`), and even the
addressing helpers (`parseRef`, `columnToIndex`, …). Branch on `.code`, never on message text.
`XlsxError extends Error`, so an existing `catch (e) { if (e instanceof Error) … }` keeps working.
Reads are guarded against decompression bombs **by default** —
a 2 GiB per-part output ceiling plus a 300× compression-ratio cap (over an 8 MiB floor); raise or
disable either via `ReadOptions.maxPartBytes` / `maxCompressionRatio` (`Number.POSITIVE_INFINITY`
to turn one off).
Every fully consumed ZIP entry must also match its declared uncompressed size. CRC-32 verification
defaults off; set `ReadOptions.verifyChecksums: true` to detect accidental corruption in fully read
entries. A cancelled stream makes no checksum claim for its partial entry. CRC-32 is not
authentication or tamper resistance.

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
`code: 'invalid-input'`. Cell and legacy Note text do not fail solely because XML cannot carry a
character: controls and lone surrogates store via Excel's `_xHHHH_` escape and round-trip. The
round trip is lossless for values, types, sheet names/order, styles, formulas, comments, pictures,
custom themes, geometry, merges, hyperlinks, visibility, defined names, tables, data
validations, conditional formatting, autofilter ranges, the 1904 date system, protection (sheet/workbook locks +
per-cell locked/hidden, password hashes verbatim), and print setup (margins, page setup,
print options, header/footer). Documented drops (never silent): in-cell rich text flattens to
plain text; an autofilter's per-column criteria/sort, row/column outline grouping, sheet tab
colors, document properties, pivot tables, external-workbook links, gradient fills, and threaded
-comment threading are not carried; and **VBA macros** — an `.xlsm` reads but rewrites to a plain
`.xlsx` without them (`Workbook.macroEnabled` flags a macro-enabled source). Full list: the root
README's fidelity table.

Worksheet tab names, legacy Note authors, `Hyperlink.location` values, and every formula field —
cell formulas, `DefinedName.refersTo`, data-validation `formula1`/`formula2`, conditional-formatting
`formulas`, and table `totalsRowFormula`/`calculatedColumnFormula` (`ST_Formula` ⊂ `ST_Xstring`) —
use OOXML's `ST_Xstring` collision rule: logical `_x0041_` stores as `_x005F_x0041_`, and the reader
decodes each complete field exactly once. Supply every formula in logical stored form without a
leading `=` and locations as logical text; never pre-encode `_xHHHH_`. Without `target`, a location
is an in-workbook reference such as `'Sheet2'!B5`; with `target`, it is a fragment within that
external destination.

Files written by openjsxl 1.0.0 with a raw escape-shaped value in these fields are ambiguous—Excel
reads `_x0041_` as `A`, and no reader can recover whether the producer intended the literal
spelling. openjsxl follows OOXML/Excel without a heuristic legacy mode. Separately, since 1.0.1 a carriage return inside any formula field travels as the `ST_Xstring` escape `_x000D_` (not `&#13;`): the 1.0.1 reader reads both spellings, but an openjsxl reader older than 1.0.1 shows the seven literal characters `_x000D_` for a CR written by 1.0.1. Direct writer locations
must be XML-safe. On read, a location that decodes to a value rejected by the XML writer is omitted;
if no target remains, the hyperlink is omitted. openpyxl 3.1.5 does not consistently apply this
rule to these schema fields.

## Other formats (read-only)

openjsxl writes `.xlsx`, but reads more: `openXlsb` (Excel Binary Workbook), `openOds`
(OpenDocument), and `openCsv` (delimited text) return the SAME `Workbook` as `openXlsx`, and
`detectSpreadsheetFormat(bytes)` → `'xlsx' | 'xlsb' | 'ods' | 'csv' | undefined` routes by content
(container formats by their package; CSV by a documented text heuristic). Accessors a format can't
express degrade (`style()` → `undefined`, `mergedCells` → `[]`), never throw. `.xlsb`/`.ods` carry
values, dates (an `.xlsb` on the epoch its `BrtWbProp` 1904 flag selects), merges (ods) and
hyperlinks; `.csv` infers numbers & booleans only (never dates).
Any of them converts to `.xlsx` through the bridge (`workbookToInput` → `writeXlsx`).
The XLSB reader supports full and compact cell records emitted by SheetJS 0.20.3, including
plain inline strings and shared strings, with positions preserved across blanks and gaps.
Rich inline `BrtCellRString` values remain unsupported; their coordinates still position
following compact cells correctly.

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
- **Names, protection & print setup (1.0):** `DefinedName`, `SheetAutoFilter`, `SheetProtection`,
  `WorkbookProtection`, `CellProtection`, `PageMargins`, `PageSetup`, `PrintOptions`, `HeaderFooter`
- **A1 & dates:** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef`, `serialToDate`,
  `dateToSerial`
- **Formulas (opt-in, `@openjsxl/core/formula`):** `parseFormula`, `evaluateWorkbook`,
  `evaluateCell`, `FormulaError` (+ `FormulaErrorCode`), `EvaluateOptions`, `FunctionSpec`
  (+ `EagerFunctionSpec`/`LazyFunctionSpec`/`EvalContext`/`ArgThunk`), `EvalValue`/`ScalarValue`,
  `FormulaErrorValue`/`ErrorCode`/`errorValue`/`isErrorValue`, `RangeView`/`isRangeView`, and the
  AST node types. A separate entry point — importing it never changes the core `"."` bundle. ~90
  built-in functions; register your own through `options.functions`.

Full guide, the complete **API reference**, design notes, and roadmap:
<https://github.com/joaquimserafim/openjsxl>

## License

MIT
