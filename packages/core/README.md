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

**ESM-only** (no CommonJS build) and, on Node, **Node ≥ 24** — a deliberate support choice, so there
is one Node line to test against. The platform Web APIs openjsxl uses
(`DecompressionStream`/`CompressionStream` with `'deflate-raw'`, `TextEncoder`/`TextDecoder`) exist
well before Node 24 (from 20.12 on the 20 line), but only Node ≥ 24 is tested and supported. The
`engines` field is advisory: npm warns on an older Node and installs anyway. `import` is the entry
point; `require('@openjsxl/core')` also resolves on Node ≥ 24 via `require(esm)`.

**Security:** see the [security
policy](https://github.com/joaquimserafim/openjsxl/blob/main/SECURITY.md) for private vulnerability
reporting, supported releases and limits when processing untrusted files. It notes which protections
are new in 1.1.0.

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
// bytes decompress on first call, each image part once per workbook)
sheet.tables; // [{ name, ref, columns, headerRow, totalsRow, style? }, …] — defined tables (0.9)
sheet.dataValidations; // [{ sqref, type, operator?, formula1?, … }, …] — dropdowns/input rules (0.9)
sheet.conditionalFormatting; // [{ sqref, rules }, …] — highlight/colorScale/dataBar/iconSet (0.9)
wb.resolveColor({ theme: 4, tint: 0.4 }); // "FF96B4D8" — theme color → ARGB, per this file's theme
wb.date1904; // false — true for a workbook on the legacy 1904 date system (since 1.1.0)

// Constant-memory streaming for large sheets — one row at a time.
for await (const row of streamSheetRows(await readFile('huge.xlsx'))) {
	console.log(row.index, row.cells.length);
}
```

`streamSheetRows` yields rows as they decode, so a damaged sheet can throw after some or all of its
rows: a sheet entry that ends short, or one whose CRC-32 is wrong when `verifyChecksums` is on,
fails only after the last row. Treat streamed rows as provisional until the loop completes.

Malformed input throws a typed `XlsxError` with a discriminating `.code` (`'not-a-zip' | 'not-xlsx'
| 'missing-part' | 'corrupt-zip' | 'part-too-large' | …`), never a bare `TypeError` from a corrupt
file. This is a uniform contract: **every public function reports failure by throwing `XlsxError`**
— the readers, the writer (`'invalid-input'`), and even the addressing helpers (`parseRef`,
`columnToIndex`, …). Branch on `.code`, never on message text. `XlsxError extends Error`, so an
existing `catch (e) { if (e instanceof Error) … }` keeps working. Reads are guarded against
decompression bombs **by default** — a 2 GiB per-part output ceiling plus a 300× compression-ratio
cap (over an 8 MiB floor); raise or disable either via `ReadOptions.maxPartBytes` /
`maxCompressionRatio` (`Number.POSITIVE_INFINITY` to turn one off). Since 1.1.0,
`ReadOptions.maxTotalBytes` can also bound the decompressed bytes of the whole archive; it is
unlimited by default. Since 1.1.0, too, each of these limits must be a number, at least 0, and
`verifyChecksums` a boolean: `null` keeps the default, and any other value, such as `NaN`, fails
with `XlsxError('invalid-input')` before the file is read. Every fully consumed ZIP entry must match
its declared size: one that ends short fails with `corrupt-zip`, and one that runs past it fails
with `part-too-large`, even when you set no limit. CRC-32 verification defaults off; set
`ReadOptions.verifyChecksums: true` (since 1.1.0) to detect accidental corruption in fully read
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

Cells can carry styles (`{ value, style }` — the same shape `style(ref)` returns) or formulas (`{
formula, value? }` — the cached value is what non-recalculating readers see), and sheets take
`columns` (widths), `rowProperties` (heights), `freeze`, `merges`, `hyperlinks`, `comments` (written
with the legacy VML part Excel needs to display them), a visibility `state`, `images` (the same `{
anchor, bytes, mime, name? }` records `images()` returns — png, jpeg, gif, bmp, tiff, webp, emf,
wmf; identical bytes dedupe into one media part), and `tables` / `dataValidations` /
`conditionalFormatting` (0.9). For huge exports, `streamXlsx` accepts the same input shape with each
sheet's `rows` as any sync/async iterable and returns a `ReadableStream` — roughly constant memory
at any row count (constant in *rows*; embedded image bytes are held, by reference, until the media
parts flush at stream end). `writeXlsx(input, { date1904: true })` selects the legacy 1904 date
epoch. Dates are UTC wall-clock values, read and written alike: build a `Date` with `Date.UTC(…)` or
an ISO string that is a bare date or ends in `Z`, and read a date cell with the `getUTC*` methods.
In the 1900 system, serials below 61 are one day off from Excel, which counts a 1900-02-29 that
never existed. Output is deterministic; unrepresentable input (no sheets, bad/duplicate sheet name,
non-finite number, invalid `Date`, malformed or overlapping merges) throws `XlsxError` with `code:
'invalid-input'`. Since 1.1.0 the writer reads only an input object's own properties: a class
instance or `Object.create` record whose prototype supplies a model field fails with `invalid-input`
naming the field, and a key added to `Object.prototype` is ignored. Cell and legacy Note text do not
fail solely because XML cannot carry a character: controls and lone surrogates store via Excel's
`_xHHHH_` escape and round-trip. The round trip is lossless for values, types, sheet names/order,
styles, formulas, comments, pictures, custom themes, geometry, merges, hyperlinks, visibility,
defined names, tables, data validations, conditional formatting, autofilter ranges, the 1904 date
system, protection (sheet/workbook locks + per-cell locked/hidden, password hashes verbatim), and
print setup (margins, page setup, print options, header/footer). Documented drops (never silent):
in-cell rich text flattens to plain text; an autofilter's per-column criteria/sort, row/column
outline grouping, sheet tab colors, document properties, pivot tables, external-workbook links,
gradient fills, and threaded -comment threading are not carried; and **VBA macros** — an `.xlsm`
reads but rewrites to a plain `.xlsx` without them (`Workbook.macroEnabled` flags a macro-enabled
source). An Excel 4.0 (XLM) macro sheet is carried as a regular worksheet with its cells and cached
values; its macro semantics are dropped, as for VBA. Full list: the root README's fidelity table.

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
spelling. openjsxl follows OOXML/Excel without a heuristic legacy mode.

Separately, since 1.1.0 a carriage return in a formula field, a hyperlink location or a Note author
is written as the `ST_Xstring` escape `_x000D_`, not the character reference `&#13;`. Readers differ
on it. openjsxl since 1.1.0 reads both spellings as a CR. openjsxl 1.0.0 and openpyxl 3.1.5 show the
seven characters `_x000D_` literally in all three: formula fields, hyperlink locations and Note
authors. SheetJS decodes it in formulas and defined names, but shows it literally in Note authors
and in a hyperlink's `location`.

Direct writer locations must be XML-safe. On read, a location that decodes to a value rejected by the
XML writer is omitted. Since 1.1.0 the `.xlsx` reader leaves out every hyperlink left with no
destination: no external target it can resolve and no usable location, such as a tooltip-only link
or one whose `r:id` names no relationship. 1.0.0 returned such a record. openpyxl 3.1.5 does not
consistently apply this rule to these schema fields, so its behavior can differ for this case.

A defined name is left out when its decoded `refersTo` is one the writer could not emit: empty,
longer than 8,192 characters, starting with `=`, or holding a control character other than tab, line
feed or carriage return, U+FFFE, U+FFFF or a lone surrogate. Since 1.1.0 this covers an `_xHHHH_`
escape that decodes to such a character, which 1.0.0 returned as the escape's text.

In a workbook that `openXlsx` or `streamSheetRows` reads (`.xlsx`, `.xlsm`, `.xltx`), two tabs can
carry the same name ignoring case once decoded: `A` and `a`, or a raw `_x0041_` and `A` (openpyxl
3.1.5 writes a tab titled `_x0041_` raw, and it decodes to `A`). Since 1.1.0 the reader keeps the
first of them that loads and leaves each later one out, together with its sheet-scoped defined names;
1.0.0 returned both. A formula or defined name that refers to a left-out tab, such as `a!A1`, then
reads the kept tab in the formula evaluator, where sheet names match ignoring case. Names are
compared with JavaScript's `toLowerCase`, the rule `writeXlsx` uses to refuse two sheet names that
match ignoring case: the Kelvin sign `K` (U+212A) matches `K`, and `İ` matches `i̇` (`i` and
U+0307), while `ß` and `SS` stay apart. Excel's comparison of non-ASCII names may differ.

## Other formats (read-only)

openjsxl writes `.xlsx`, but reads more: `openXlsb` (Excel Binary Workbook), `openOds`
(OpenDocument), and `openCsv` (delimited text) return the SAME `Workbook` as `openXlsx`, and
`detectSpreadsheetFormat(bytes)` → `'xlsx' | 'xlsb' | 'ods' | 'csv' | undefined` routes by content
(container formats by their package; CSV by a documented text heuristic). Accessors a format can't
express degrade (`style()` → `undefined`, `mergedCells` → `[]`), never throw. `.xlsb` carries
values, dates (on the epoch its `BrtWbProp` 1904 flag selects) and external hyperlink targets, but
not a link's in-workbook location, tooltip or display text; an empty `.xlsb` sheet has no
`dimension`, where `.xlsx` declares `A1`. `.ods` carries values, dates, merges and in-cell
hyperlinks (an in-document `#Sheet.A1` link reads as an external `target`). `.csv` infers numbers &
booleans only (never dates): it drops a leading `+` (`+123` reads as 123) and rounds a decimal with
more digits than a double holds, while an integer beyond `Number.MAX_SAFE_INTEGER` or with leading
zeros stays text. Any of them converts to `.xlsx` through the bridge (`workbookToInput` →
`writeXlsx`). The XLSB reader supports full and compact cell records emitted by SheetJS 0.20.3,
including plain inline strings and shared strings, with positions preserved across blanks and gaps.
Rich inline `BrtCellRString` values remain unsupported; their coordinates still position following
compact cells correctly.

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
  built-in functions; register your own through `options.functions`. The root README's Formulas
  section lists the 1.1.0 formula changes and what is not evaluated.

Full guide, the complete **API reference**, and design notes:
<https://github.com/joaquimserafim/openjsxl>

## License

MIT
