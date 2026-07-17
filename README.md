# openjsxl

[![npm version](https://img.shields.io/npm/v/openjsxl?color=cb3837&logo=npm)](https://www.npmjs.com/package/openjsxl)
[![CI](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml/badge.svg)](https://github.com/joaquimserafim/openjsxl/actions/workflows/ci.yml)
[![install size](https://packagephobia.com/badge?p=openjsxl)](https://packagephobia.com/result?p=openjsxl)
[![types included](https://img.shields.io/npm/types/openjsxl)](https://www.npmjs.com/package/openjsxl)
[![license: MIT](https://img.shields.io/npm/l/openjsxl?color=blue)](./LICENSE)

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **reader + writer — pre-1.0.** Read typed cells, styles, number formats, formulas,
> merged ranges, hyperlinks, comments, and pictures; stream large sheets in roughly constant
> memory; get typed errors on malformed input. Write `.xlsx` from plain data with `writeXlsx`,
> and read → modify → write with `workbookToInput`. **New in 0.6:** anchored pictures — read
> them with `sheet.images()`, write them via `images` on a sheet, and they round-trip through
> the bridge byte-exact (png, jpeg, gif, bmp, tiff, webp, emf, wmf). **New in 0.7:** it now
> *reads* more than it writes — `.xlsb`, `.ods`, and `.csv`/`.tsv` open into the same `Workbook`,
> with `detectSpreadsheetFormat` to route by content. **New in 0.8:** the opt-in `openjsxl/formula`
> entry *evaluates* formulas — `evaluateWorkbook`/`evaluateCell` over 90+ built-ins, plus your own.
> **New in 0.9:** three structural features that read AND write — defined **tables**, **data
> validation** (dropdowns and input rules), and **conditional formatting** (highlights, colour scales,
> data bars, icon sets) — plus hardening: string content survives XML-illegal characters via
> ST_Xstring escaping, and every read is guarded against decompression bombs by default (a 2 GiB
> per-part ceiling + a 300× compression-ratio cap, both tunable).
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
	// Zip-bomb guards are ON by default (a 2 GiB per-part ceiling + a 300× compression-ratio
	// cap); tighten or disable them per read — here, a stricter 50 MB ceiling.
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
sheet.formula('E1'); // "B1*2" — the formula text (shared formulas come back translated)
sheet.dimension; // "A1:E2" | undefined (the declared used range)
sheet.mergedCells; // ["A1:B1", "A2:A4", …]
sheet.hyperlinks; // [{ ref: "A1", target: "https://…", tooltip?, location?, display? }, …]
sheet.comments; // [{ ref: "A1", author: "Ada", text: "note" }, …]
sheet.columns; // [{ min: 2, max: 3, width: 25.5, hidden? }, …] — column geometry
sheet.rowProperties; // Map<row, { height?, hidden? }>
sheet.freeze; // { rows?, cols? } | undefined — the frozen pane
sheet.state; // "visible" | "hidden" | "veryHidden" (sheet.visible is the boolean)

await sheet.images(); // [{ anchor, bytes, mime, name? }, …] — anchored pictures; `bytes` is the
// raw image payload, `anchor` the raw cell + EMU geometry. Async & lazy: media is only
// decompressed on first call (once per drawing part — sheets sharing a logo each read it once)

wb.resolveColor({ theme: 4, tint: 0.4 }); // "FF96B4D8" — a style color as 8-digit ARGB,
// resolved against the workbook's own theme (colors stay raw {theme, tint} in the style model)
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
name, a non-finite number, an invalid `Date` — throws a typed `XlsxError` with
`code: 'invalid-input'` rather than producing a file Excel must repair. String *content* never
throws: characters XML can't carry (control chars, lone surrogates) are stored with the same
`_xHHHH_` escape Excel uses, so they round-trip instead of corrupting or rejecting.
An optional second argument `writeXlsx(input, { date1904: true })` selects the 1904 date epoch
(legacy Mac workbooks); dates are 1900-epoch by default.

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

### Comments, formulas & streaming (0.5)

A cell can carry a formula (`{ formula, value? }` — the cached `value` is what non-recalculating
readers see), sheets take `comments` (written with the legacy VML part Excel needs to actually
*show* them on hover), and `streamXlsx` is the constant-memory mirror of `writeXlsx` — each
sheet's `rows` may be any sync or async iterable (a DB cursor, a paged API), pulled only as the
output stream drains, so the full sheet never lives in memory:

```ts
import { streamXlsx, writeXlsx } from 'openjsxl';

const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [[42, { formula: 'A1*2', value: 84 }]],
			comments: [{ ref: 'A1', author: 'Ada', text: 'the answer' }],
		},
	],
});

// Same input shape, but rows can be lazy — memory stays flat at any row count.
const stream = streamXlsx({ sheets: [{ name: 'Big', rows: millionRowCursor() }] });
await stream.pipeTo(destination);
```

The streaming writer's flat memory is **constant in rows** — row data is pulled and released as
the output drains. Pictures are the exception: image bytes (passed by reference, never copied)
are held until the media parts flush at the end of the stream.

### Pictures (0.6)

Sheets take `images` — the same `{ anchor, bytes, mime, name? }` records `sheet.images()`
returns, so pictures pass straight through a round trip. The anchor is raw OOXML geometry: a
1-based `from` cell plus sizes in EMU (≈ 9 525 EMU per pixel at 96 dpi). Pin a fixed-size logo
with `ext`, or span cells with `to`:

```ts
const bytes = await writeXlsx({
	sheets: [
		{
			name: 'Report',
			rows: [['Q3 Sales']],
			images: [
				{
					anchor: { from: { col: 4, row: 1 }, ext: { cx: 96 * 9525, cy: 96 * 9525 } },
					bytes: logoBytes, // Uint8Array — written verbatim, never decoded
					mime: 'image/png', // png, jpeg, gif, bmp, tiff, webp, emf, wmf
					name: 'Logo',
				},
			],
		},
	],
});
```

Identical image bytes are deduplicated into one media part workbook-wide, and a workbook
without images emits no drawing machinery at all.

### Read → modify → write

`workbookToInput` turns an open `Workbook` back into writer input, so you can round-trip a file:

```ts
import { openXlsx, workbookToInput, writeXlsx } from 'openjsxl';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await openXlsx(await readFile('in.xlsx'));
const input = await workbookToInput(wb);
const sheets = input.sheets.map((sheet, i) =>
	i === 0 ? { ...sheet, rows: [...sheet.rows, ['appended', 'row']] } : sheet,
); // tweak the plain data (the input is deeply readonly — spread, don't mutate)
await writeFile('out.xlsx', await writeXlsx({ ...input, sheets }));
```

The round trip is **lossless for values, types, sheet names/order, styles, formulas, comments,
pictures, geometry, structural metadata, defined names, tables, data validation, conditional
formatting, autofilters, protection, and print setup**:

| Round-trips losslessly | Not carried (yet) |
| --- | --- |
| string, number, boolean, `Date` values | error cells without a formula (written as their text) |
| number formats — built-in & custom codes | absolute-anchored pictures & non-picture drawings (shapes, charts) — skipped on read |
| fonts, fills, borders, alignment | picture effects (crop, rotation, borders) — a picture carries anchor + bytes + type + name |
| colors: rgb, indexed, theme + tint (raw) | in-cell rich-text runs — flattened to plain text (values survive; per-run bold/color is lost) |
| control chars & `_xHHHH_` literals in strings (ST_Xstring escape) | autofilter **criteria** & sort state — the filter range carries, the per-column filter/sort is dropped |
| custom theme part (carried byte-identical) | printer-settings binary (`printerSettings.bin`) & manual page breaks — dropped |
| formula text + cached value | **VBA macros** — an `.xlsm` opens and reads, but rewriting writes a plain `.xlsx` and the macros are dropped (check `Workbook.macroEnabled` to warn first) |
| defined names / named ranges (global & sheet-scoped, incl. `_xlnm.*` built-ins) | row/column **outline grouping** (`outlineLevel`/`collapsed`) — the group nesting is dropped (values, width/height/hidden survive) |
| comments (author + text, Excel-visible) | sheet **tab colors** and other `sheetPr` — dropped |
| anchored pictures (bytes byte-exact, anchor, type, name) | **document properties** (author, title, created/modified) — no `docProps` is emitted (deterministic bytes) |
| empty cells (sparse), incl. styled blanks | **pivot tables** — dropped |
| column widths, row heights, hidden, freeze | **external workbook links** — a formula's `[1]Sheet!` reference re-emits as text with no target |
| merged ranges & hyperlinks | **gradient fills** — read as no fill |
| sheet visibility (hidden / veryHidden) | **threaded comment** thread structure — the text survives as a legacy comment; replies/authorship threading is dropped |
| sheet names & tab order | `calcChain`/`calcPr` — dropped; Excel recomputes the dependency chain on open |
| tables (name, range, columns, header/totals, style) | |
| data validations (dropdowns & input rules) | |
| conditional formatting (highlights, scales, bars, icon sets) | |
| autofilter range (filter dropdowns; the paired `_xlnm._FilterDatabase` is managed automatically) | |
| protection — sheet & workbook locks, per-cell locked/hidden, password hashes carried verbatim | |
| print setup — margins, orientation, scale, fit-to-page, paper size, print options, header/footer | |

Documented flattenings (values stay exact; internal spelling normalizes): row/column *default*
styles resolve into per-cell styles (each cell keeps its effective format); shared and array
formulas re-emit as per-cell plain formulas (the same text Excel shows in each cell — data-table
formulas keep only their cached value); a tolerated non-canonical cell ref spelling (e.g.
lowercase `a1`) re-emits canonically (`A1`); a foreign producer's out-of-spec table name (a space,
a cell-reference shape, empty) normalizes into a legal identifier so the table re-saves rather than
aborting; and media parts renumber with alternate extension spellings normalized (`tif` → `tiff`,
`jpg` → `jpeg` — the image bytes themselves are untouched).

## Reading other formats (0.7)

openjsxl **writes `.xlsx`**, but it **reads more**: `openXlsb` (Excel Binary Workbook), `openOds`
(OpenDocument / LibreOffice), and `openCsv` (delimited text) all return the SAME `Workbook` as
`openXlsx` — typed cells, the same accessors, the same discriminated-union `Cell`.
`detectSpreadsheetFormat(bytes)` sniffs the container so "a user uploaded a spreadsheet" is one
code path, and any reader becomes a converter to `.xlsx` through the bridge:

```ts
import {
	detectSpreadsheetFormat, openXlsx, openXlsb, openOds, openCsv,
	workbookToInput, writeXlsx,
} from 'openjsxl';
import { readFile, writeFile } from 'node:fs/promises';

const bytes = await readFile('upload.bin');
const wb = await (async () => {
	switch (await detectSpreadsheetFormat(bytes)) {
		case 'xlsx': return openXlsx(bytes); // also .xlsm / .xltx / .xltm
		case 'xlsb': return openXlsb(bytes);
		case 'ods':  return openOds(bytes);
		case 'csv':  return openCsv(bytes);  // synchronous — CSV has no container
		default: throw new Error('unrecognized spreadsheet format');
	}
})();

await writeFile('out.xlsx', await writeXlsx(await workbookToInput(wb))); // convert to .xlsx
```

Each reader returns the shared model; features a format can't express **degrade** (`style()` →
`undefined`, `mergedCells` → `[]`, `images()` → `[]`), never throw. What each format carries:

| Format | Reads | Not carried (degrades) |
| --- | --- | --- |
| **`.xlsx` / `.xlsm` / `.xltx`** | everything in the round-trip table above | — |
| **`.xlsb`** | values & types, dates (style-driven), number formats, hyperlinks, dimension, sheet visibility | merges, formula text (cached values kept), styles, comments, geometry, images |
| **`.ods`** | values & types, dates, merges, in-cell hyperlinks, sheet names/order | styles & number formats, formula text (cached values kept), visibility, comments, geometry, images |
| **`.csv` / `.tsv`** | values with conservative type inference (numbers & booleans; **never dates**), one sheet, auto-detected delimiter | everything structural — CSV carries none |

`detectSpreadsheetFormat` returns `'xlsx' | 'xlsb' | 'ods' | 'csv' | undefined`. Container formats
are identified by their package; **CSV has no magic bytes**, so any non-zip input that decodes as
UTF-8 text classifies as `'csv'` — a documented best-effort heuristic, not a guarantee. Encrypted
`.ods` and other unreadable inputs fail with a typed `XlsxError`.

## Formulas (0.8)

The reader keeps a formula's **text** and its cached value; the opt-in `openjsxl/formula` entry adds
a zero-dependency **engine** that recomputes them. It's a separate import so a consumer who never
touches formulas never loads a byte of the parser or evaluator, and the core `openjsxl` bundle is
byte-for-byte unchanged whether or not it exists.

```ts
import { openXlsx } from 'openjsxl';
import { evaluateWorkbook, evaluateCell, parseFormula } from 'openjsxl/formula';

const wb = await openXlsx(bytes);
const result = await evaluateWorkbook(wb);   // every formula cell, recomputed
result.get('Sheet1', 'B2');                  // e.g. 84  (a number | string | boolean | error value)

await evaluateCell(wb, 'Sheet1', 'B2');      // or just one cell (+ its dependencies)
parseFormula('SUM(A1:A9)*2');                // …or just the typed AST, no evaluation
```

Evaluation is **read-only** — the workbook and its stored `<v>` caches are never mutated, so a
recomputed value can (correctly) supersede a stale cache. Deep dependency chains don't grow the JS
stack, and **circular references resolve to a dedicated `#CYCLE!` value** rather than hanging — the
cycle's cells become `#CYCLE!` while every unrelated cell still evaluates.

**Bring your own functions.** `options.functions` registers user-defined functions with the same
shape the 90+ built-ins use (case-insensitive, eager or lazy):

```ts
await evaluateWorkbook(wb, {
  functions: { MILESTOKM: { minArgs: 1, maxArgs: 1, evaluate: ([mi]) => Number(mi) * 1.60934 } },
});
```

**Determinism is enforced.** Volatile functions have no ambient source: `TODAY`/`NOW` and
`RAND`/`RANDBETWEEN` throw a typed `FormulaError` unless you inject `options.now` / `options.random`,
so the same inputs always produce the same output.

**Built-in coverage (90+).** Math/stats (`SUM` `AVERAGE` `COUNT*` `MIN` `MAX` `MEDIAN` `LARGE`
`SMALL` `ROUND*` `INT` `ABS` `MOD` `POWER` `SQRT` `EXP` `LN` `LOG*` `SIGN` `TRUNC` `CEILING` `FLOOR`
`SUMPRODUCT`), logical (`IF` `IFERROR` `IFNA` `IFS` `SWITCH` `AND` `OR` `XOR` `NOT`), lookup
(`VLOOKUP` `HLOOKUP` `INDEX` `MATCH` `CHOOSE` `ROWS` `COLUMNS`), conditional aggregates (`SUMIF(S)`
`COUNTIF(S)` `AVERAGEIF(S)`), text (`CONCAT` `TEXTJOIN` `LEN` `LEFT` `RIGHT` `MID` `TRIM` `UPPER`
`LOWER` `PROPER` `SUBSTITUTE` `REPLACE` `FIND` `SEARCH` `VALUE` `REPT` `EXACT` `CHAR` `CODE`),
information (`IS*` `N` `T` `NA` `ERROR.TYPE`), and date/time (`DATE` `YEAR` `MONTH` `DAY` `HOUR`
`MINUTE` `SECOND` `WEEKDAY` `TIME` `DAYS` `EDATE` `EOMONTH`).

**Not evaluated in 0.8** (documented, never silently wrong — an unknown function is `#NAME?`, an
unsupported reference a typed error value): `ROW`/`COLUMN`, `XLOOKUP`, `DATEVALUE`, `TEXT`, wildcard
`SEARCH` (matched literally), array arithmetic inside an aggregator (`SUM(A1:A3*B1:B3)` → a typed
`#VALUE!`, not an element-wise array), 3-D references (`Sheet1:Sheet3!A1` → `#REF!`),
`OFFSET`/`INDIRECT`, R1C1, cross-workbook references, and iterative-calculation mode. Semantics
follow documented Excel; a single-cell reference passed to an aggregate is treated as a literal (so
text/booleans in one referenced cell coerce rather than being ignored — multi-cell ranges are
Excel-exact).

## Tables, data validation & conditional formatting (0.9)

Three structural features that read AND write, each as the same record on both sides — so
read → modify → write carries them across untouched:

- **Tables** (`Worksheet.tables` ↔ `SheetInput.tables`) — a defined table's name, range, columns
  (with totals-row label/function and formulas), header/totals flags, and style banding. On write,
  column names derive from the header row; the numeric id and part number are assigned for you.
- **Data validation** (`Worksheet.dataValidations` ↔ `SheetInput.dataValidations`) — dropdowns and
  input rules: all eight types (`list`/`whole`/`decimal`/`date`/`time`/`textLength`/`custom`/`none`),
  the comparison operator, operand text, prompt/error messages, and `showDropDown` in its intuitive
  sense (`true` = arrow shown; the file's inverted attribute is translated for you).
- **Conditional formatting** (`Worksheet.conditionalFormatting` ↔ `SheetInput.conditionalFormatting`)
  — highlight rules (`cellIs`/`expression`/`top10`/text/… with an inline differential style), color
  scales, data bars, and icon sets. Rule priorities are renumbered densely on write.
- **AutoFilter** (`Worksheet.autoFilter` ↔ `SheetInput.autoFilter`) — the filter-dropdown range,
  `{ ref: 'A1:C10' }`. On write the paired hidden `_xlnm._FilterDatabase` defined name Excel expects
  is created for you; per-column filter *criteria* and sort state are not carried.
- **Protection** (`Worksheet.protection` ↔ `SheetInput.protection`, `Workbook.protection` ↔
  `WorkbookInput.protection`, `CellStyle.protection`) — lock a sheet (`{ sheet: true }`) or the workbook
  structure (`{ lockStructure: true }`), and mark cells `{ locked, hidden }` (cells are locked by default;
  set `locked: false` to leave a cell editable in a protected sheet). Any password hash is carried
  verbatim — openjsxl never computes, verifies, or strips one.

```ts
const bytes = await writeXlsx({
	sheets: [{
		name: 'Inventory',
		rows: [['Item', 'Qty', 'Status'], ['Pears', 8, 'Low']],
		tables: [{ name: 'Stock', ref: 'A1:C2', columns: [], headerRow: true, totalsRow: false }],
		dataValidations: [{ sqref: ['C2'], type: 'list', formula1: '"In stock,Low,Out"', showDropDown: true }],
		conditionalFormatting: [{
			sqref: ['B2'],
			rules: [{ type: 'cellIs', priority: 1, operator: 'lessThan', formulas: ['10'],
				dxf: { fill: { bgColor: { rgb: 'FFFFC7CE' } } } }],
		}],
		autoFilter: { ref: 'A1:C2' },
	}],
});
```

Consistent with the whole library, the tolerant reader normalizes a foreign producer's out-of-spec
table (an odd name, a mismatched column count, a totals row with nowhere to go) into something the
strict writer accepts, so a table file from another tool re-saves instead of aborting.

## Defined names, protection & print setup (1.0)

The workbook-fidelity features that complete a faithful round trip — all read AND write, each the
same record on both sides:

- **Defined names** (`Workbook.definedNames` ↔ `WorkbookInput.definedNames`) — named ranges and
  constants, workbook-global or sheet-scoped (`localSheetId`); `refersTo` is stored form (no leading `=`).
- **Protection** (`Worksheet.protection` ↔ `SheetInput.protection`, `Workbook.protection` ↔
  `WorkbookInput.protection`, `CellStyle.protection`) — lock a sheet (`{ sheet: true }`) or the workbook
  structure (`{ lockStructure: true }`); cells are locked by default, so `{ locked: false }` leaves a cell
  editable under protection. Any password hash is carried **verbatim** — openjsxl never computes, verifies,
  or strips one.
- **Print setup** (`Worksheet.pageMargins`/`pageSetup`/`printOptions`/`headerFooter` ↔ the same
  `SheetInput` keys) — margins (inches), orientation / scale / fit-to-page / paper size, gridline &
  heading printing, and `&`-code headers/footers.

```ts
const bytes = await writeXlsx({
	definedNames: [{ name: 'RevenueRange', refersTo: 'Q3!$B$2:$B$4' }],
	protection: { lockStructure: true }, // lock workbook structure
	sheets: [{
		name: 'Q3',
		rows: [['Region', 'Rep', 'Revenue'], ['EMEA', 'Ada', 4200]],
		autoFilter: { ref: 'A1:C2' },
		protection: { sheet: true }, // protect the sheet (cells are locked by default)
		pageSetup: { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 },
		headerFooter: { oddHeader: '&CQ3 Revenue', oddFooter: '&RPage &P of &N' },
	}],
});
```

An `.xlsm` opens and reads, but rewriting through the bridge writes a plain `.xlsx` and drops the VBA
project — check `Workbook.macroEnabled` to warn a user before a rewrite discards their macros. See
[`examples/14`](./examples/14-names-autofilter-protection-print.mjs) for the full round trip.

## API reference

The complete public surface, verified against the built type declarations. openjsxl has two
entry points: **`openjsxl`** (reading, writing, addressing) and the opt-in **`openjsxl/formula`**
(a formula parser + evaluator that the core bundle never loads unless you import it). Every
function reports failure by throwing a typed [`XlsxError`](#errors) — branch on `.code`, never on
message text. Each section links to a runnable script in [`examples/`](./examples).

### `openjsxl` — functions

**Reading** — every reader returns the same [`Workbook`](#workbook); a format's unsupported
features degrade (empty/`undefined`) rather than throwing.
([01](./examples/01-read-cells.mjs)–[04](./examples/04-metadata.mjs),
[11](./examples/11-other-formats.mjs))

| Function | Signature | Notes |
| --- | --- | --- |
| `openXlsx` | `(source: Uint8Array \| ArrayBuffer, options?: ReadOptions) => Promise<Workbook>` | Open `.xlsx`/`.xlsm`/`.xltx`/`.xltm`. Random-access `cell()`. |
| `streamSheetRows` | `(source, sheetName?: string, options?: ReadOptions) => AsyncGenerator<Row>` | Constant-memory row stream; `sheetName` defaults to the first sheet. |
| `openXlsb` | `(source, options?: ReadOptions) => Promise<Workbook>` | Excel Binary Workbook (`.xlsb`). |
| `openOds` | `(source, options?: ReadOptions) => Promise<Workbook>` | OpenDocument spreadsheet (`.ods`). |
| `openCsv` | `(source: Uint8Array \| ArrayBuffer \| string, options?: CsvReadOptions) => Workbook` | Delimited text. **Synchronous** — CSV has no container to decompress. |
| `detectSpreadsheetFormat` | `(source, options?: ReadOptions) => Promise<SpreadsheetFormat \| undefined>` | Sniff bytes → `'xlsx' \| 'xlsb' \| 'ods' \| 'csv'`, or `undefined`. |

**Writing** — describe a workbook as plain data; cell types are inferred from the JS values.
([06](./examples/06-write.mjs)–[10](./examples/10-images.mjs),
[13](./examples/13-tables-validation-formatting.mjs),
[14](./examples/14-names-autofilter-protection-print.mjs))

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
| `serialToDate` | `(serial: number, date1904?: boolean) => Date` | Excel serial → `Date`. |
| `dateToSerial` | `(date: Date, date1904?: boolean) => number` | `Date` → Excel serial. |

### `Workbook`

Returned by every reader. ([01](./examples/01-read-cells.mjs),
[09](./examples/09-comments-formulas-theme.mjs),
[14](./examples/14-names-autofilter-protection-print.mjs))

| Member | Type | Description |
| --- | --- | --- |
| `sheets` | `readonly SheetInfo[]` | Sheets in tab order. |
| `definedNames` | `readonly DefinedName[]` | Workbook defined (named) ranges/constants; `[]` for ods/xlsb/csv. |
| `protection` | `WorkbookProtection \| undefined` | Workbook-level `<workbookProtection>`. |
| `macroEnabled` | `boolean` | `true` for a read `.xlsm`/`.xltm`. Rewriting drops the VBA project. |
| `themeXml` (getter) | `string \| undefined` | Raw `theme1.xml`, or `undefined` when absent. |
| `sheet(name)` | `(name: string) => Worksheet` | The worksheet with this tab name. Throws if none. |
| `resolveColor(color)` | `(color: Color) => string \| undefined` | Resolve a raw `Color` to 8-digit ARGB (`undefined` for auto/indexed/unresolved theme). |

### `Worksheet`

A sheet's cells, style/format accessors, and metadata.
([01](./examples/01-read-cells.mjs), [04](./examples/04-metadata.mjs))

| Member | Type | Description |
| --- | --- | --- |
| `name` / `path` | `string` | Tab name / part locator. |
| `visible` / `state` | `boolean` / `SheetState` | Visibility (`state` is the source; `visible === (state === "visible")`). |
| `mergedCells` | `readonly string[]` | Merged ranges in A1, document order. |
| `hyperlinks` | `readonly Hyperlink[]` | Hyperlinks, document order. |
| `dimension` | `string \| undefined` | Declared used range in A1. |
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
| `formula(ref)` | `(ref: string) => string \| undefined` | Formula text at a cell. |
| `images()` | `() => Promise<readonly SheetImage[]>` | Anchored pictures (lazy). |
| `cell(ref)` | `(ref: string) => Cell` | Cell at A1; absent cells read as `empty`. |
| `rows()` | `() => AsyncGenerator<Row>` | Stream populated rows (sparse). |

### `Cell` & `Row`

`Cell` is a discriminated union — narrow on `type` to type `value`. `Row` is `{ index: number;
cells: readonly Cell[] }` (sparse — absent rows/cells are omitted).

| `cell.type` | `cell.value` | | `cell.type` | `cell.value` |
| --- | --- | --- | --- | --- |
| `"empty"` | `null` | | `"boolean"` | `boolean` |
| `"string"` | `string` | | `"date"` | `Date` |
| `"number"` | `number` | | `"error"` | `string` |

### Style & geometry types

The style model `style(ref)` returns and `{ value, style }` accepts. ([07](./examples/07-styles-and-layout.mjs))

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

([13](./examples/13-tables-validation-formatting.mjs))

- **`TableInfo`** `{ name, ref, columns: TableColumn[], headerRow, totalsRow, style?: TableStyleInfo, headerRowStyle?, dataStyle?, totalsRowStyle? }`; **`TableColumn`**, **`TableStyleInfo`**.
- **`DataValidation`** `{ sqref: string[], type: DataValidationType, operator?, formula1?, formula2?, allowBlank?, showDropDown?, showInputMessage?, showErrorMessage?, errorStyle?, promptTitle?, prompt?, errorTitle?, error? }`; enums `DataValidationType`, `DataValidationOperator`, `DataValidationErrorStyle`. `showDropDown` is intuitive (`true` = arrow shown).
- **`ConditionalFormatting`** `{ sqref: string[], rules: ConditionalFormattingRule[] }`; `ConditionalFormattingRule` = `CfHighlightRule` | `CfColorScaleRule` | `CfDataBarRule` | `CfIconSetRule` (discriminated by `type`); `CfHighlightType`, `Cfvo`.
- **`DxfStyle`** `{ numberFormat?, font?, fill?: DxfFill, border?, alignment? }` — a differential highlight; `DxfFill` is kept raw (visible color is `bgColor`).

### Writer input types

What `writeXlsx` / `streamXlsx` accept — mirrors of the reader's model, so read → modify → write
is a pass-through. ([06](./examples/06-write.mjs), [08](./examples/08-streaming-write.mjs))

- **`WorkbookInput`** `{ sheets: SheetInput[], themeXml?, definedNames?: DefinedName[], protection?: WorkbookProtection }`.
- **`SheetInput`** `{ name, rows: (CellInput[] | undefined)[], columns?, rowProperties?, freeze?, merges?, hyperlinks?, state?, comments?, images?, tables?, dataValidations?, conditionalFormatting?, autoFilter?, protection?, pageMargins?, pageSetup?, printOptions?, headerFooter? }`.
- **`CellInput`** = `CellValue | StyledCell`. **`CellValue`** = `string | number | boolean | Date | null | undefined`. **`StyledCell`** `{ value?, style?: CellStyle, formula? }`.
- **`WriteOptions`** `{ date1904? }`.
- **Streaming** — **`StreamWorkbookInput`** `{ sheets: StreamSheetInput[], themeXml?, definedNames?, protection? }`; **`StreamSheetInput`** (as `SheetInput` but `rows: StreamRows`); **`StreamRows`** = `Iterable | AsyncIterable` of `CellInput[] | undefined`.

### Names, protection & print setup types

The 1.0 fidelity types — each is the same record on the reader and the writer.
([14](./examples/14-names-autofilter-protection-print.mjs))

- **`DefinedName`** `{ name, refersTo, localSheetId?, hidden? }` — a named range/constant (`refersTo` is stored form, no leading `=`).
- **`SheetAutoFilter`** `{ ref }` — the filter range; the paired `_xlnm._FilterDatabase` is managed for you.
- **`SheetProtection`** — `{ sheet?, objects?, scenarios?, formatCells?, …, password?, algorithmName?, hashValue?, saltValue?, spinCount? }`. Password material carried verbatim (never computed).
- **`WorkbookProtection`** — `{ lockStructure?, lockWindows?, workbookPassword?, workbookAlgorithmName?, workbookHashValue?, workbookSaltValue?, workbookSpinCount? }`.
- **`CellProtection`** `{ locked?, hidden? }` — on `CellStyle.protection`; only meaningful under sheet protection.
- **`PageMargins`** `{ left, right, top, bottom, header, footer }` (inches; all six required).
- **`PageSetup`** `{ paperSize?, orientation?, scale?, fitToWidth?, fitToHeight?, firstPageNumber?, useFirstPageNumber?, blackAndWhite?, draft?, cellComments?, pageOrder? }`.
- **`PrintOptions`** `{ gridLines?, headings?, horizontalCentered?, verticalCentered? }`.
- **`HeaderFooter`** `{ oddHeader?, oddFooter?, evenHeader?, evenFooter?, firstHeader?, firstFooter?, differentOddEven?, differentFirst?, scaleWithDoc?, alignWithMargins? }` (Excel `&`-codes, verbatim).

### Errors

([05](./examples/05-error-handling.mjs))

- **`XlsxError extends Error`** — `{ code: XlsxErrorCode, message, cause? }`. The single error type every public function throws.
- **`XlsxErrorCode`** = `"not-a-zip"` | `"not-xlsx"` | `"missing-part"` | `"corrupt-zip"` | `"unsupported"` | `"no-such-sheet"` | `"part-too-large"` | `"invalid-input"`.
- **`ReadOptions`** `{ maxPartBytes?, maxCompressionRatio? }` — zip-bomb guards, on by default (2 GiB per-part ceiling; 300× ratio cap). `CsvReadOptions` `{ delimiter?, sheetName?, inferTypes? }`; `SpreadsheetFormat`.

### `openjsxl/formula`

The opt-in evaluator — a separate import that never changes the core bundle.
([12](./examples/12-formulas.mjs))

| Function | Signature | Notes |
| --- | --- | --- |
| `parseFormula` | `(text: string) => FormulaAst` | Stored-form text → typed AST. Throws `FormulaError`. |
| `evaluateWorkbook` | `(workbook: Workbook, options?: EvaluateOptions) => Promise<WorkbookEvalResult>` | Recompute every formula cell (read-only). |
| `evaluateCell` | `(workbook: Workbook, sheet: string, ref: string, options?: EvaluateOptions) => Promise<ScalarValue>` | One cell + its dependencies. |
| `errorValue` | `(code: ErrorCode) => FormulaErrorValue` | The interned error value for a code. |
| `isErrorValue` | `(v: unknown) => v is FormulaErrorValue` | Narrowing guard. |
| `isRangeView` | `(v: unknown) => v is RangeView` | Narrowing guard. |

- **`EvaluateOptions`** `{ functions?: Record<string, unknown>, now?: () => Date, random?: () => number, maxCellVisits? }` — inject a clock/RNG for volatile functions (else they throw), register UDFs (each value a `FunctionSpec`, validated at runtime), cap evaluations.
- **Results** — `WorkbookEvalResult` `{ cells: readonly CellResult[]; get(sheet, ref): ScalarValue | undefined }`; `CellResult` `{ sheet, ref, value }`.
- **Values** — `ScalarValue` = `number | string | boolean | null | FormulaErrorValue`; `EvalValue` = `ScalarValue | RangeView`; `FormulaErrorValue` `{ kind: "error", code: ErrorCode }`; `ErrorCode` (the eight `ST_CellErrorType` values + `#CYCLE!`).
- **`RangeView`** — a lazy window over a reference's used cells: `sheet`, `startCol/startRow/endCol/endRow`, `width`/`height`/`cellCount` (getters), `entries()`, `values()`, `populatedCount()`, `single()`, `topLeft()`, `cellAt(rowOffset, colOffset)`; `RangeEntry` `{ col, row, value }`.
- **User functions** — `FunctionSpec` = `EagerFunctionSpec | LazyFunctionSpec`, both extending `FunctionSpecBase` `{ minArgs, maxArgs, volatile? }`; `evaluate(args, ctx: EvalContext)` where `EvalContext` `{ now(): number; random(): number }` and (lazy) `ArgThunk` = `() => EvalValue`.
- **AST** — `FormulaAst` is the union of `NumberLiteral`, `StringLiteral`, `BooleanLiteral`, `ErrorLiteral`, `ArrayLiteral` (`ArrayElement`), `CellRefNode`, `NameRef`, `RangeRef`, `UnionRef`, `FunctionCall`, `EmptyArg`, `UnaryOp`, `PostfixOp`, `BinaryOp`, `Group`, `StructuredRef`, `ExternalRef`; a qualifying sheet is `SheetSpec`. The single-cell node is **`CellRefNode`** (named so it never collides with `openjsxl`'s `CellRef`, letting you import both entry points at once).
- **Errors** — `FormulaError extends Error` `{ code: FormulaErrorCode, position? }`; `FormulaErrorCode` = `"parse-error"` | `"depth-exceeded"` | `"too-many-args"` | `"budget-exceeded"` | `"volatile-unconfigured"` | `"unsupported"`.

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
| **read** | **0.70 s · 203 MB** | 1.6 s · 806 MB | 2.2 s · 543 MB |
| **write** | **0.75 s · 393 MB** | 3.2 s · 1.4 GB | 1.9 s · 564 MB |

Writing with `streamXlsx` from a lazy row source holds memory roughly **flat (~100 MB)** no matter
the row count (flat in *rows* — embedded images, when present, stay resident until the stream ends).

**Reads more than `.xlsx`.** The same million cells as `.xlsb`, `.ods`, and `.csv`, each parsed and
materialized cell-by-cell — openjsxl leads every format (e.g. `.xlsb` in **0.19 s** vs SheetJS's
1.63 s; ExcelJS reads neither `.xlsb` nor `.ods`).

**And it stays small** — a clean production install (`npm install --omit=dev`), the library plus every
runtime dependency on disk:

| | openjsxl | ExcelJS `4.4.0` | SheetJS `0.18.5` |
| --- | --- | --- | --- |
| runtime deps | **0 third-party** (1 pkg — its own core) | 96 packages | 8 packages |
| installed | **~0.45 MB** | 34 MB | 14 MB |

The full matrix (10k / 100k / 1M cells, numbers / strings / styled, read + write), the four-format
read lanes, the library-size table, the methodology, and out-of-band **openpyxl** /
**python-calamine** reference numbers are in [`docs/benchmarks.md`](./docs/benchmarks.md) — reproduce
it end-to-end with `pnpm bench`. _(Measured 2026-07-17; every "fast" in these docs traces back here.)_

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
