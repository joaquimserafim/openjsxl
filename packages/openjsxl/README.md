# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

Fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) reader **and writer** for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge. This is the package to install; it re-exports
the [`@openjsxl/core`](https://www.npmjs.com/package/@openjsxl/core) engine.

On the 2026-10-22 benchmark, this workspace build reads a 1M-cell numeric XLSX sheet
in **805 ms** (ExcelJS: 1.69 s; SheetJS: 2.62 s).
Published 1.0.0 installs in **0.52 MB with zero third-party dependencies**. XLSB reads now pass
the expected checksum at all three benchmark sizes; the 1M-cell XLSB read takes **990 ms**.
See the [full measurements and limitations](https://github.com/joaquimserafim/openjsxl/blob/main/docs/benchmarks.md).

```sh
npm install openjsxl
```

**Security:** see the [security
policy](https://github.com/joaquimserafim/openjsxl/blob/main/SECURITY.md) for private vulnerability
reporting, supported releases and limits when processing untrusted files. It notes which protections
are new in 1.1.0.

### Requirements

**ESM-only** (there is no CommonJS build) and, on Node, **Node ≥ 24** — a deliberate support choice,
so there is one Node line to test against. The platform Web APIs openjsxl uses
(`DecompressionStream`/`CompressionStream` with `'deflate-raw'`, `TextEncoder`/`TextDecoder`) exist
well before Node 24 (from 20.12 on the 20 line), but only Node ≥ 24 is tested and supported. The
`engines` field is advisory: npm warns on an older Node and installs anyway. `import` is the entry
point; `require('openjsxl')` also resolves on Node ≥ 24 via `require(esm)`. Deno, Bun, browsers, and
edge runtimes need no special setup.

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

For very large sheets use `streamSheetRows` (constant memory). It yields rows as they decode, so a
damaged sheet can throw after some or all of its rows: a sheet entry that ends short, or one whose
CRC-32 is wrong when `verifyChecksums` is on, fails only after the last row. Treat streamed rows as
provisional until the loop completes. A worksheet also exposes `style(ref)`, `numberFormat`,
`formula(ref)`, `dimension`, `mergedCells`, `hyperlinks`, `comments`, `columns`, `rowProperties`,
`freeze`, `state`/`visible`, async `images()` (anchored pictures — raw bytes, media type, and cell +
EMU anchor), and `tables`, `dataValidations`, `conditionalFormatting` (0.9); the workbook resolves
theme colors to ARGB with `resolveColor`; and the reader throws a typed `XlsxError` (with a
discriminating `code`) on malformed input — including hostile input: every read is guarded against
decompression bombs by default (a 2 GiB per-part ceiling + a 300× compression-ratio cap, tunable via
`ReadOptions`). Every fully consumed ZIP entry must match its declared size: one that ends short
fails with `corrupt-zip`, and one that runs past it fails with `part-too-large`, even when you set
no limit. Set `ReadOptions.verifyChecksums: true` (since 1.1.0) to opt into CRC-32 verification (off
by default); cancelled streams make no claim for partially read entries. CRC-32 detects accidental
corruption, not intentional tampering.

**Writing:** describe a workbook as plain data and get back `.xlsx` bytes — cell types are inferred
from the JS values. Cells can carry styles (`{ value, style }` — the same shape `style(ref)`
returns) or formulas (`{ formula, value? }`), and sheets take column widths, row heights, frozen
panes, merged ranges, hyperlinks, comments (Excel-visible), a visibility state, `images` (the same
records `images()` returns — png, jpeg, gif, bmp, tiff, webp, emf, wmf), and `tables` /
`dataValidations` / `conditionalFormatting` (0.9). For exports too big to buffer, `streamXlsx`
streams the same input shape from lazy (sync or async) row sources with roughly constant memory —
constant in *rows*; embedded image bytes are held until the media parts flush at stream end:

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

Since 1.1.0 the writer reads only an input object's own properties: a class instance or
`Object.create` record whose prototype supplies a model field fails with `invalid-input` naming the
field, and a key added to `Object.prototype` is ignored.

Dates are UTC wall-clock values, read and written alike: build a `Date` with `Date.UTC(…)` or an
ISO string that is a bare date or ends in `Z`, and read a date cell with the `getUTC*` methods. In
the 1900 system, serials below 61 are one day off from Excel, which counts a 1900-02-29 that never
existed.

`workbookToInput` turns an open `Workbook` back into writer input for read → modify → write —
values, types, styles, formulas, comments, pictures, custom themes, geometry, merges, hyperlinks,
sheet visibility, defined names, tables, data validations, conditional formatting, autofilter
ranges, the 1904 date system, protection (sheet/workbook locks + per-cell locked/hidden), and print
setup (margins, orientation, scale, header/footer) all round-trip. Documented drops (never silent):
in-cell rich text flattens to plain text; autofilter criteria/sort, outline grouping, tab colors,
document properties, pivot tables, external links, and gradient fills are not carried; and **VBA
macros** — an `.xlsm` reads but rewrites without them (`Workbook.macroEnabled` flags it). An Excel
4.0 (XLM) macro sheet is carried as a regular worksheet with its cells and cached values; its macro
semantics are dropped, as for VBA. Full list: the project README's fidelity table.

Worksheet tab names, legacy Note authors, `Hyperlink.location` values, and **every formula field** —
cell formulas, `DefinedName.refersTo`, data-validation `formula1`/`formula2`, conditional-formatting
`formulas`, and table `totalsRowFormula`/`calculatedColumnFormula` (`ST_Formula` is a restriction of
`ST_Xstring`) — use OOXML's `ST_Xstring` collision rule: a logical literal `_x0041_` stores as
`_x005F_x0041_`, and the reader decodes each complete field exactly once. Supply every formula in
logical stored form without a leading `=` and locations as logical text; never pre-encode `_xHHHH_`.
Without `target`, a location is an in-workbook reference such as `'Sheet2'!B5`; with `target`, it is
a fragment within that external destination.

Compatibility note: openjsxl 1.0.0 wrote these fields raw, so an escape-shaped value in a
1.0.0-produced file is ambiguous and Microsoft Excel interprets it as an escape (`_x0041_` → `A`).
There is no reliable way to infer the original intent; openjsxl follows OOXML/Excel rather than
double-decoding or guessing.

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

**Reading other formats (0.7):** openjsxl writes `.xlsx` but reads more — `openXlsb` (Excel Binary
Workbook), `openOds` (OpenDocument), and `openCsv` (delimited text) all return the SAME `Workbook`,
and `detectSpreadsheetFormat(bytes)` routes by content (`'xlsx' | 'xlsb' | 'ods' | 'csv' |
undefined`; `.xlsm`/`.xltx` read as `'xlsx'`). Accessors a format can't express degrade, never
throw. `.xlsb` carries values, dates (on the epoch its `BrtWbProp` 1904 flag selects) and external
hyperlink targets, but not a link's in-workbook location, tooltip or display text; an empty `.xlsb`
sheet has no `dimension`, where `.xlsx` declares `A1`. `.ods` carries values, dates, merges and
in-cell hyperlinks (an in-document `#Sheet.A1` link reads as an external `target`). `.csv` infers
numbers & booleans only (never dates): it drops a leading `+` (`+123` reads as 123) and rounds a
decimal with more digits than a double holds, while an integer beyond `Number.MAX_SAFE_INTEGER` or
with leading zeros stays text. So "a user uploaded a spreadsheet" is one code path, and any reader
converts to `.xlsx` through the bridge. The XLSB reader supports full and compact cell records
emitted by SheetJS 0.20.3, including plain inline strings and shared strings, with positions
preserved across blanks and gaps. Rich inline `BrtCellRString` values remain unsupported; their
coordinates still position following compact cells correctly.

**Formulas (0.8):** the reader keeps a formula's text and cached value; the opt-in
`openjsxl/formula` entry adds a zero-dependency engine that recomputes them — `evaluateWorkbook(wb)`
/ `evaluateCell(wb, sheet, ref)` over 90+ built-in functions, plus your own via `options.functions`.
It's a separate import (the core bundle is unchanged whether or not you use it); evaluation is
read-only (it can supersede a stale cache), a circular reference gives a `#CYCLE!` value instead of
hanging (an `IFERROR` around it can return its fallback, and a formula past the depth cap can abort
the evaluation), and volatile functions (`TODAY`/`RAND`) require an injected clock/RNG so results
stay deterministic. Since 1.1.0, a call Excel stores as `_xlfn.NAME` reaches `NAME`, a sheet name in
a formula matches ignoring case, a single-cell reference reads as a reference in most built-ins
(`SUM(A1)` ignores text in A1), an array constant is a full array where a function takes a range,
`&`, `SUBSTITUTE`, `REPLACE`, `UPPER` and `LOWER` cap the text they build at 32,767 characters, and
a criterion or text lookup value over 255 characters gives `#VALUE!`. The project README's Formulas
section has the details, and lists what is not evaluated.

**Tables, data validation & conditional formatting (0.9):** three structural features that read
AND write, each as the same record on both sides. `tables` (name, range, columns, header/totals
flags, style banding), `dataValidations` (dropdowns and input rules — all eight types, operators,
prompt/error text, intuitive `showDropDown`), and `conditionalFormatting` (highlight rules with an
inline differential style, color scales, data bars, icon sets). They round-trip through
`workbookToInput`, and the tolerant reader normalizes a foreign producer's out-of-spec table into
something the strict writer accepts, so a table file from another tool re-saves instead of aborting.

See the [project README](https://github.com/joaquimserafim/openjsxl#readme) for the full guide and
design notes.

## API reference

The complete public surface, verified against the built type declarations. openjsxl has two
entry points: **`openjsxl`** (reading, writing, addressing) and the opt-in **`openjsxl/formula`**
(a formula parser + evaluator that the core bundle never loads unless you import it). Every
function reports failure by throwing a typed [`XlsxError`](#errors) — branch on `.code`, never on
message text. Each section links to a runnable script in [`examples/`](https://github.com/joaquimserafim/openjsxl/tree/main/examples).

### `openjsxl` — functions

**Reading** — every reader returns the same [`Workbook`](#workbook); a format's unsupported
features degrade (empty/`undefined`) rather than throwing.
([01](https://github.com/joaquimserafim/openjsxl/blob/main/examples/01-read-cells.mjs)–[04](https://github.com/joaquimserafim/openjsxl/blob/main/examples/04-metadata.mjs),
[11](https://github.com/joaquimserafim/openjsxl/blob/main/examples/11-other-formats.mjs))

| Function | Signature | Notes |
| --- | --- | --- |
| `openXlsx` | `(source: Uint8Array \| ArrayBuffer, options?: ReadOptions) => Promise<Workbook>` | Open `.xlsx`/`.xlsm`/`.xltx`/`.xltm`. Random-access `cell()`. |
| `streamSheetRows` | `(source, sheetName?: string, options?: ReadOptions) => AsyncGenerator<Row>` | Constant-memory row stream; `sheetName` defaults to the first sheet. |
| `openXlsb` | `(source, options?: ReadOptions) => Promise<Workbook>` | Excel Binary Workbook (`.xlsb`). |
| `openOds` | `(source, options?: ReadOptions) => Promise<Workbook>` | OpenDocument spreadsheet (`.ods`). |
| `openCsv` | `(source: Uint8Array \| ArrayBuffer \| string, options?: CsvReadOptions) => Workbook` | Delimited text. **Synchronous** — CSV has no container to decompress. |
| `detectSpreadsheetFormat` | `(source, options?: ReadOptions) => Promise<SpreadsheetFormat \| undefined>` | Sniff bytes → `'xlsx' \| 'xlsb' \| 'ods' \| 'csv'`, or `undefined`. |

**Writing** — describe a workbook as plain data; cell types are inferred from the JS values.
([06](https://github.com/joaquimserafim/openjsxl/blob/main/examples/06-write.mjs)–[10](https://github.com/joaquimserafim/openjsxl/blob/main/examples/10-images.mjs),
[13](https://github.com/joaquimserafim/openjsxl/blob/main/examples/13-tables-validation-formatting.mjs),
[14](https://github.com/joaquimserafim/openjsxl/blob/main/examples/14-names-autofilter-protection-print.mjs))

| Function | Signature | Notes |
| --- | --- | --- |
| `writeXlsx` | `(workbook: WorkbookInput, options?: WriteOptions) => Promise<Uint8Array>` | `.xlsx` bytes. Throws `invalid-input` for anything unrepresentable. |
| `streamXlsx` | `(workbook: StreamWorkbookInput, options?: WriteOptions) => ReadableStream<Uint8Array>` | Constant-memory mirror of `writeXlsx`; invalid input surfaces on the stream. |
| `workbookToInput` | `(workbook: Workbook) => Promise<WorkbookInput>` | Turn an open workbook back into writer input (read → modify → write). |

**Addressing & dates** — pure helpers. The four A1 helpers throw `XlsxError('invalid-input')` on
a malformed reference (a uniform contract, since 1.0).

| Function | Signature | Notes |
| --- | --- | --- |
| `columnToIndex` | `(letters: string) => number` | `"A"` → `1`, `"AA"` → `27`. Throws on a non-letter or overflowing ref. |
| `indexToColumn` | `(index: number) => string` | `1` → `"A"`. Throws on a non-integer or `< 1`. |
| `parseRef` | `(ref: string) => CellRef` | `"B3"` → `{ col: 2, row: 3 }`. Throws on a non-A1 string. |
| `formatRef` | `(ref: CellRef) => string` | `{ col: 2, row: 3 }` → `"B3"`. Throws on a non-integer/`< 1` row. |
| `serialToDate` | `(serial: number, date1904?: boolean) => Date` | Excel serial → UTC `Date`. |
| `dateToSerial` | `(date: Date, date1904?: boolean) => number` | UTC `Date` → Excel serial. |

### `Workbook`

Returned by every reader. ([01](https://github.com/joaquimserafim/openjsxl/blob/main/examples/01-read-cells.mjs),
[09](https://github.com/joaquimserafim/openjsxl/blob/main/examples/09-comments-formulas-theme.mjs),
[14](https://github.com/joaquimserafim/openjsxl/blob/main/examples/14-names-autofilter-protection-print.mjs))

| Member | Type | Description |
| --- | --- | --- |
| `sheets` | `readonly SheetInfo[]` | Sheets in tab order. |
| `definedNames` | `readonly DefinedName[]` | Workbook defined (named) ranges/constants; `[]` for ods/xlsb/csv. |
| `protection` | `WorkbookProtection \| undefined` | Workbook-level `<workbookProtection>`. |
| `macroEnabled` | `boolean` | `true` for a read `.xlsm`/`.xltm`. Rewriting drops the VBA project. |
| `date1904` | `boolean` | Since 1.1.0. The workbook's date system: `true` for the legacy 1904 epoch (`<workbookPr date1904>`), else `false` (`.xlsb` reads its `BrtWbProp` flag; always `false` for ods/csv). The formula engine uses it for every serial ↔ date conversion, and the bridge carries it. |
| `themeXml` (getter) | `string \| undefined` | Raw `theme1.xml`, or `undefined` when absent. |
| `sheet(name)` | `(name: string) => Worksheet` | The worksheet with this tab name. Throws if none. |
| `resolveColor(color)` | `(color: Color) => string \| undefined` | Resolve a raw `Color` to 8-digit ARGB (`undefined` for auto/indexed/unresolved theme). |

### `Worksheet`

A sheet's cells, style/format accessors, and metadata.
([01](https://github.com/joaquimserafim/openjsxl/blob/main/examples/01-read-cells.mjs), [04](https://github.com/joaquimserafim/openjsxl/blob/main/examples/04-metadata.mjs))

| Member | Type | Description |
| --- | --- | --- |
| `name` / `path` | `string` | Tab name / part locator. |
| `visible` / `state` | `boolean` / `SheetState` | Visibility (`state` is the source; `visible === (state === "visible")`). |
| `mergedCells` | `readonly string[]` | Merged ranges in A1, document order. |
| `hyperlinks` | `readonly Hyperlink[]` | Hyperlinks in document order; location is in-workbook without target, or a fragment within target when both are present. |
| `dimension` | `string \| undefined` | Declared used range in A1. An empty `.xlsb` sheet has none, where `.xlsx` declares `A1`. |
| `comments` | `readonly Comment[]` | Cell comments. |
| `tables` | `readonly TableInfo[]` | Defined tables. |
| `dataValidations` | `readonly DataValidation[]` | Validation rules. |
| `conditionalFormatting` | `readonly ConditionalFormatting[]` | CF blocks. |
| `autoFilter` | `SheetAutoFilter \| undefined` | Filter-dropdown range. |
| `protection` | `SheetProtection \| undefined` | `<sheetProtection>`. |
| `pageMargins` / `pageSetup` / `printOptions` / `headerFooter` | `… \| undefined` | Print setup. |
| `columns` | `readonly ColumnProps[]` | Column width/visibility. |
| `rowProperties` | `ReadonlyMap<number, RowProps>` | Per-row height/visibility. |
| `freeze` | `FreezePane \| undefined` | Frozen pane. |
| `numberFormat(ref)` | `(ref: string) => string \| undefined` | Format code at a cell. |
| `style(ref)` | `(ref: string) => CellStyle \| undefined` | Resolved style at a cell. |
| `formula(ref)` | `(ref: string) => string \| undefined` | Formula text at a cell — logical stored form (`ST_Xstring`-decoded once). |
| `images()` | `() => Promise<readonly SheetImage[]>` | Anchored pictures (lazy). |
| `cell(ref)` | `(ref: string) => Cell` | Cell at A1; absent cells read as `empty`. |
| `rows()` | `() => AsyncGenerator<Row>` | Stream rows (sparse; in `.xlsx`, a property-only row and a styled blank cell are included). |

### `Cell` & `Row`

`Cell` is a discriminated union — narrow on `type` to type `value`. `Row` is `{ index: number;
cells: readonly Cell[] }` (sparse — absent rows/cells are omitted; in `.xlsx`, a row with only a
height or hidden flag, and a styled blank cell, are present).

| `cell.type` | `cell.value` | | `cell.type` | `cell.value` |
| --- | --- | --- | --- | --- |
| `"empty"` | `null` | | `"boolean"` | `boolean` |
| `"string"` | `string` | | `"date"` | `Date` |
| `"number"` | `number` | | `"error"` | `string` |

A `"date"` value is a UTC wall-clock `Date`: read it with `getUTC*`.

### Style & geometry types

The style model `style(ref)` returns and `{ value, style }` accepts. ([07](https://github.com/joaquimserafim/openjsxl/blob/main/examples/07-styles-and-layout.mjs))

- **`CellStyle`** — `{ numberFormat?, font?: FontStyle, fill?: FillStyle, border?: BorderStyle, alignment?: Alignment, protection?: CellProtection }`. Every component optional; the workbook default resolves to `undefined`.
- **`FontStyle`** — `{ name?, size?, bold?, italic?, underline?: UnderlineStyle, strike?, color?: Color }`.
- **`FillStyle`** — `{ patternType: PatternType, fgColor?: Color, bgColor?: Color }` (solid uses `fgColor`).
- **`BorderStyle`** / **`BorderEdge`** — per-edge `{ top?, right?, bottom?, left?: BorderEdge }`, each `{ style: BorderLineStyle, color?: Color }`.
- **`Alignment`** — `{ horizontal?, vertical?, wrapText?, shrinkToFit?, indent?, textRotation? }`.
- **`Color`** — raw, never resolved: `{ rgb }` | `{ theme, tint? }` | `{ indexed }` | `{ auto: true }`.
- **Enums** — `CellType`, `PatternType`, `BorderLineStyle`, `UnderlineStyle` (`"single"`/`"double"`), `HorizontalAlignment`, `VerticalAlignment`.
- **Geometry** — `ColumnProps` `{ min, max, width?, hidden? }`, `RowProps` `{ height?, hidden? }`, `FreezePane` `{ rows?, cols? }`, `SheetImage` `{ anchor: ImageAnchor, bytes, mime, name? }`, `ImageAnchor` `{ from: AnchorPoint, to?, ext?, editAs? }`, `AnchorPoint` `{ col, row, colOff?, rowOff? }`.
- **Metadata** — `SheetInfo` `{ name, path, visible, state }`, `SheetState`, `Comment` `{ ref, author?, text }`, `Hyperlink` `{ ref, target?, location?, tooltip?, display? }`.

### Tables, validation & conditional formatting types

([13](https://github.com/joaquimserafim/openjsxl/blob/main/examples/13-tables-validation-formatting.mjs))

- **`TableInfo`** `{ name, ref, columns: TableColumn[], headerRow, totalsRow, style?: TableStyleInfo, headerRowStyle?, dataStyle?, totalsRowStyle? }`; **`TableColumn`**, **`TableStyleInfo`**.
- **`DataValidation`** `{ sqref: string[], type: DataValidationType, operator?, formula1?, formula2?, allowBlank?, showDropDown?, showInputMessage?, showErrorMessage?, errorStyle?, promptTitle?, prompt?, errorTitle?, error? }`; enums `DataValidationType`, `DataValidationOperator`, `DataValidationErrorStyle`. `showDropDown` is intuitive (`true` = arrow shown).
- **`ConditionalFormatting`** `{ sqref: string[], rules: ConditionalFormattingRule[] }`; `ConditionalFormattingRule` = `CfHighlightRule` | `CfColorScaleRule` | `CfDataBarRule` | `CfIconSetRule` (discriminated by `type`); `CfHighlightType`, `Cfvo`.
- **`DxfStyle`** `{ numberFormat?, font?, fill?: DxfFill, border?, alignment? }` — a differential highlight; `DxfFill` is kept raw (visible color is `bgColor`).

### Writer input types

What `writeXlsx` / `streamXlsx` accept — mirrors of the reader's model, so read → modify → write
is a pass-through. ([06](https://github.com/joaquimserafim/openjsxl/blob/main/examples/06-write.mjs), [08](https://github.com/joaquimserafim/openjsxl/blob/main/examples/08-streaming-write.mjs))

- **`WorkbookInput`** `{ sheets: SheetInput[], themeXml?, definedNames?: DefinedName[], protection?: WorkbookProtection, date1904? }` (`date1904?` since 1.1.0; a set `WriteOptions.date1904` overrides it).
- **`SheetInput`** `{ name, rows: (CellInput[] | undefined)[], columns?, rowProperties?, freeze?, merges?, hyperlinks?, state?, comments?, images?, tables?, dataValidations?, conditionalFormatting?, autoFilter?, protection?, pageMargins?, pageSetup?, printOptions?, headerFooter? }`.
- **`CellInput`** = `CellValue | StyledCell`. **`CellValue`** = `string | number | boolean | Date | null | undefined`. **`StyledCell`** `{ value?, style?: CellStyle, formula? }`.
- **`WriteOptions`** `{ date1904? }`.
- **Streaming** — **`StreamWorkbookInput`** `{ sheets: StreamSheetInput[], themeXml?, definedNames?, protection?, date1904? }` (`date1904?` since 1.1.0); **`StreamSheetInput`** (as `SheetInput` but `rows: StreamRows`); **`StreamRows`** = `Iterable | AsyncIterable` of `CellInput[] | undefined`.

### Names, protection & print setup types

The 1.0 fidelity types — each is the same record on the reader and the writer.
([14](https://github.com/joaquimserafim/openjsxl/blob/main/examples/14-names-autofilter-protection-print.mjs))

- **`DefinedName`** `{ name, refersTo, localSheetId?, hidden? }` — a named range/constant (`refersTo`
  is logical stored form without a leading `=`; do not pre-encode `_xHHHH_` text).
- **`SheetAutoFilter`** `{ ref }` — the filter range; the paired `_xlnm._FilterDatabase` is managed for you.
- **`SheetProtection`** — `{ sheet?, objects?, scenarios?, formatCells?, …, password?, algorithmName?, hashValue?, saltValue?, spinCount? }`. Password material carried verbatim (never computed).
- **`WorkbookProtection`** — `{ lockStructure?, lockWindows?, workbookPassword?, workbookAlgorithmName?, workbookHashValue?, workbookSaltValue?, workbookSpinCount? }`.
- **`CellProtection`** `{ locked?, hidden? }` — on `CellStyle.protection`; only meaningful under sheet protection.
- **`PageMargins`** `{ left, right, top, bottom, header, footer }` (inches; all six required).
- **`PageSetup`** `{ paperSize?, orientation?, scale?, fitToWidth?, fitToHeight?, firstPageNumber?, useFirstPageNumber?, blackAndWhite?, draft?, cellComments?, pageOrder? }`.
- **`PrintOptions`** `{ gridLines?, headings?, horizontalCentered?, verticalCentered? }`.
- **`HeaderFooter`** `{ oddHeader?, oddFooter?, evenHeader?, evenFooter?, firstHeader?, firstFooter?, differentOddEven?, differentFirst?, scaleWithDoc?, alignWithMargins? }` (Excel `&`-codes, verbatim).

### Errors

([05](https://github.com/joaquimserafim/openjsxl/blob/main/examples/05-error-handling.mjs))

- **`XlsxError extends Error`** — `{ code: XlsxErrorCode, message, cause? }`. The single error type every public function throws.
- **`XlsxErrorCode`** = `"not-a-zip"` | `"not-xlsx"` | `"missing-part"` | `"corrupt-zip"` | `"unsupported"` | `"no-such-sheet"` | `"part-too-large"` | `"invalid-input"`.
- **`ReadOptions`** `{ maxPartBytes?, maxCompressionRatio?, maxTotalBytes?, verifyChecksums? }` — zip-bomb guards are on by default (2 GiB per-part ceiling; 300× ratio cap); `maxTotalBytes` (since 1.1.0) optionally bounds the decompressed bytes of the whole archive and is unlimited by default; CRC-32 verification (`verifyChecksums`, since 1.1.0) defaults to `false`. Since 1.1.0 each limit must be a number, at least 0 (`Infinity` for none), and `verifyChecksums` a boolean; `null` keeps the default, and any other value throws `XlsxError('invalid-input')` before the file is read, from `openXlsx`, `openXlsb`, `openOds`, `streamSheetRows` and `detectSpreadsheetFormat`. `CsvReadOptions` `{ delimiter?, sheetName?, inferTypes? }`; `SpreadsheetFormat`.

### `openjsxl/formula`

The opt-in evaluator — a separate import that never changes the core bundle.
([12](https://github.com/joaquimserafim/openjsxl/blob/main/examples/12-formulas.mjs))

| Function | Signature | Notes |
| --- | --- | --- |
| `parseFormula` | `(text: string) => FormulaAst` | Stored-form text → typed AST. Throws `FormulaError`. |
| `evaluateWorkbook` | `(workbook: Workbook, options?: EvaluateOptions) => Promise<WorkbookEvalResult>` | Recompute every formula cell (read-only). |
| `evaluateCell` | `(workbook: Workbook, sheet: string, ref: string, options?: EvaluateOptions) => Promise<ScalarValue>` | One cell + its dependencies. |
| `errorValue` | `(code: ErrorCode) => FormulaErrorValue` | The interned error value for a code. |
| `isErrorValue` | `(v: unknown) => v is FormulaErrorValue` | Narrowing guard. |
| `isRangeView` | `(v: unknown) => v is RangeView` | Narrowing guard. |

- **`EvaluateOptions`** `{ functions?: Record<string, unknown>, now?: () => Date, random?: () => number, maxCellVisits? }` — inject a clock/RNG for volatile functions (else they throw), register UDFs (each value a `FunctionSpec`, validated at runtime), cap evaluations. `maxCellVisits` (default 5,000,000) must be a number, at least 0 (`Infinity` for none); since 1.1.0 any other value except `null` throws `FormulaError('unsupported')`.
- **Results** — `WorkbookEvalResult` `{ cells: readonly CellResult[]; get(sheet, ref): ScalarValue | undefined }`; `CellResult` `{ sheet, ref, value }`.
- **Values** — `ScalarValue` = `number | string | boolean | null | FormulaErrorValue`; `EvalValue` = `ScalarValue | RangeView`; `FormulaErrorValue` `{ kind: "error", code: ErrorCode }`; `ErrorCode` (the eight `ST_CellErrorType` values + `#CYCLE!`).
- **`RangeView`** — a lazy window over a reference's used cells: `sheet`, `startCol/startRow/endCol/endRow`, `width`/`height`/`cellCount` (getters), `entries()`, `values()`, `populatedCount()`, `single()`, `topLeft()`, `cellAt(rowOffset, colOffset)`; `RangeEntry` `{ col, row, value }`.
- **User functions** — `FunctionSpec` = `EagerFunctionSpec | LazyFunctionSpec`, both with `{ minArgs, maxArgs, volatile? }` (their shared base, `FunctionSpecBase`, is not exported); `evaluate(args, ctx: EvalContext)` where `EvalContext` `{ now(): number; random(): number; date1904?: boolean }` (`date1904?` since 1.1.0; the evaluator always sets it, and a hand-built context without it uses the 1900 system) and (lazy) `ArgThunk` = `() => EvalValue`.
- **AST** — `FormulaAst` is the union of `NumberLiteral`, `StringLiteral`, `BooleanLiteral`, `ErrorLiteral`, `ArrayLiteral` (`ArrayElement`), `CellRefNode`, `NameRef`, `RangeRef`, `UnionRef`, `FunctionCall`, `EmptyArg`, `UnaryOp`, `PostfixOp`, `BinaryOp`, `Group`, `StructuredRef`, `ExternalRef`; a qualifying sheet is `SheetSpec`. The single-cell node is **`CellRefNode`** (named so it never collides with `openjsxl`'s `CellRef`, letting you import both entry points at once).
- **Errors** — `FormulaError extends Error` `{ code: FormulaErrorCode, position? }`; `FormulaErrorCode` = `"parse-error"` | `"depth-exceeded"` | `"too-many-args"` | `"budget-exceeded"` | `"volatile-unconfigured"` | `"unsupported"`.

## License

MIT
