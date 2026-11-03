# Publishing

How to cut a release of openjsxl to npm. The repo is a pnpm workspace; releases go out with
`pnpm -r publish`, which handles dependency order and the `workspace:*` rewrite for us.

## What ships

| Package | npm | Notes |
| --- | --- | --- |
| `openjsxl` | ✅ public | The facade users install. Re-exports `@openjsxl/core`. |
| `@openjsxl/core` | ✅ public | The engine. Required — the facade's types re-export it. |
| `@openjsxl/fixtures` | ❌ never | `private: true`. Test corpus; also holds git-ignored Apache-2.0 files under `local/`. |
| `@openjsxl/bench`, `@openjsxl/fuzz`, `openjsxl-examples` | ❌ never | `private: true`. Benchmarks, the fuzz harness and the runnable examples. |

Both public packages carry `publishConfig.access: "public"` (scoped packages default to private
otherwise), `files: ["dist"]`, a `README.md`, a `LICENSE`, and a `prepack` hook that rebuilds
`dist` at pack time so a tarball can never ship stale output.

## One-time prerequisites

1. An npm **organization** named `openjsxl` (a scope must be owned before you can publish
   `@openjsxl/*`). Free to create at <https://www.npmjs.com/org/create>.
2. `npm login`, and confirm membership: `npm whoami`.
3. If your npm account enforces 2FA for publish, have your authenticator ready (`--otp=<code>`).

## Release steps

Run everything from the repo root.

### 1. Preflight — frozen install and green gates

```sh
pnpm install --frozen-lockfile
pnpm -r build
pnpm check                              # Biome lint + format check
pnpm typecheck
pnpm test
pnpm --filter openjsxl-examples all     # every example, against the built packages
pnpm audit --prod
```

### 2. Set the version

Set the **same** version in the workspace root and both public packages by editing the `"version"`
field in:

- `package.json` (the workspace root, which never publishes)
- `packages/core/package.json`
- `packages/openjsxl/package.json`

(Leave the private packages — `@openjsxl/fixtures`, `@openjsxl/bench`, `@openjsxl/fuzz` and the
examples — as they are; they never publish.) Keep the three in lock-step. Only core and the facade
publish, but the root version names the release the repository is at, and the facade depends on
`@openjsxl/core` via `workspace:*`, which is rewritten to that exact version on publish, so a
mismatch would ship a broken dependency range.

### 3. Dry-run — inspect the tarballs

```sh
pnpm -C packages/core     pack
pnpm -C packages/openjsxl pack

tar -tzf packages/core/openjsxl-core-*.tgz
tar -tzf packages/openjsxl/openjsxl-*.tgz

# The facade's core dependency must be the exact version you set — not workspace:* or a range.
# This must print one line; if it prints nothing, stop.
tar -xzOf packages/openjsxl/openjsxl-*.tgz package/package.json | grep '"@openjsxl/core": "<version>"'

rm -f packages/*/*.tgz   # clean up
```

Each tarball should contain **only** `package.json`, `README.md`, `LICENSE`, and `dist/*` — no
`src/`, `node_modules/`, `tsconfig`, `.tgz`, or `local/`.

### 4. Push `main` and wait for CI

Commit everything first — `pnpm publish` refuses an unclean tree (by design). Push `main`, then wait
until hosted CI and CodeQL pass on that exact commit. The release is published from that commit
only.

```sh
git status                        # must be clean
git push origin main
```

### 5. Publish from that commit

```sh
git rev-parse HEAD                # must be the commit CI and CodeQL passed
pnpm -r publish --access public   # add --otp=<code> if 2FA is on
```

`pnpm -r publish` publishes `@openjsxl/core` first, then `openjsxl` (topological order), rewrites
`workspace:*` → the version, and skips the private packages.

### 6. Verify and tag

```sh
npm view openjsxl version
npm view @openjsxl/core version

git tag v<version>                # never -f
git push origin v<version>        # never -f
```

If `v<version>` already exists, locally or on the remote, stop: a release tag names one commit for
good, so never move, re-point or force-push it. `git tag` refuses an existing name without `-f`; do
not add it.

## Versioning

- **`0.1.0`** — first public release: the hardened reader (typed cells, number formats, merges,
  hyperlinks, comments, constant-memory streaming, typed `XlsxError`). No writer yet.
- **`0.2.0`** — documentation release: self-contained per-package READMEs, a `PUBLISHING.md`
  runbook, and a runnable `examples/` workspace. Reader code unchanged from `0.1.0`.
- **`0.2.1`** — drop published source maps (install size ~178 KB → ~55 KB). No API change.
- **`0.3.0`** — the writer: `writeXlsx` (author an `.xlsx` from plain data — values, types, and
  multiple sheets) and `workbookToInput` (read → modify → write). Round-trip is lossless for
  values, types, and sheet names/order; verified against real fixtures and openpyxl. Reader
  unchanged apart from a robustness fix (an out-of-range date serial now reads as a number).
- **`0.4.0`** — styles & layout: cell styles read (`style(ref)`) and write (`{ value, style }`),
  number-format codes, sheet geometry (column widths, row heights, frozen panes), and structural
  metadata (merged ranges, hyperlinks, sheet visibility). Round-trip now also carries styles,
  geometry, merges, hyperlinks, and visibility — the drop-list is comments, formulas, and error
  cells. Additive reader API: `Worksheet.style/columns/rowProperties/freeze/state`.
- **`0.5.0`** — fidelity + streaming: comments write (with the legacy VML part Excel needs to
  show them), formula text read/translate/write (shared formulas come back per-cell translated;
  cached values kept), custom-theme carry (the theme part round-trips byte-identical) and
  `Workbook.resolveColor` (`{theme, tint}` → ARGB), and `streamXlsx` — a constant-memory
  streaming writer fed by sync/async row iterables. The round-trip drop-list is down to bare
  error cells. Additive reader API: `Worksheet.formula(ref)`, `Workbook.resolveColor`/`themeXml`.
  Published, reproducible benchmarks land in `docs/benchmarks.md` (`pnpm bench`).
- **`0.5.1`** — two writer fixes from a post-release cross-cutting review: `streamXlsx` no longer
  drops `rowProperties` addressed past the last streamed row (it now emits the trailing
  property-only `<row/>` elements exactly like `writeXlsx`), and the zip entry cap is 65 534 in
  both writers so the EOCD count can never carry `0xffff`, the ZIP64 sentinel. Benchmarks
  re-measured on the fixed build (no regression). No API change.
- **`0.6.0`** — images: anchored-picture read (async `Worksheet.images()` — raw bytes, media
  type, cell + EMU anchor, name), picture write on both writers (`images` on a sheet; identical
  bytes dedupe into one media part workbook-wide), and bridge carry — pictures round-trip
  byte-exact for the full read set (png, jpeg, gif, bmp, tiff, webp, emf, wmf). The tolerant
  reader clamps out-of-range anchor values into writable bounds, so one malformed picture can't
  make a file un-rewritable. Additive API: `Worksheet.images()`, `SheetInput.images`,
  `SheetImage`/`ImageAnchor`/`AnchorPoint` types. Round-trip drop-list: bare error cells;
  absolute-anchored and non-picture drawings (skipped on read); picture effects.
- **`0.7.0`** — more formats to read: `openXlsb` (Excel Binary Workbook / BIFF12), `openOds`
  (OpenDocument), and `openCsv` (`.csv`/`.tsv`, RFC 4180) all open into the same `Workbook` surface
  as `openXlsx`, so "a user uploaded a spreadsheet" is one code path; `detectSpreadsheetFormat`
  sniffs the container to route by content, and `.xlsm`/`.xltx` open through `openXlsx`. Read-only —
  conversion to `.xlsx` is the bridge (`workbookToInput` → `writeXlsx`), and per-format accessors a
  format can't express degrade rather than throw (the drop-list is a documented matrix). The writer
  is unchanged, so `.xlsx` output stays byte-identical. Additive API: `openXlsb`/`openOds`/`openCsv`,
  `detectSpreadsheetFormat`, `SpreadsheetFormat`/`CsvReadOptions` types, and `Worksheet`/`Row` are now
  structural interfaces (the xlsx implementation is `XlsxWorksheet` — `instanceof Worksheet`, never
  documented, no longer applies).
- **`0.8.0`** — formulas: an opt-in, zero-dependency **evaluation engine** behind the new
  `openjsxl/formula` entry point. `parseFormula` (typed AST), `evaluateWorkbook`/`evaluateCell`
  (pull-based, memoizing, cycle-safe — circular references resolve to a `#CYCLE!` value, deep chains
  don't grow the JS stack), and ~97 built-in functions (SUM/IF/VLOOKUP/INDEX/MATCH/SUMIF(S)/text/
  date/…) plus caller-registered UDFs via `options.functions`. Evaluation is read-only (it can
  supersede a stale cache), and volatile functions (`TODAY`/`RAND`) require an injected clock/RNG so
  results stay deterministic. The entry is module-graph-isolated: importing it never changes the
  core `"."` bundle, and the writer is untouched, so `.xlsx` output stays byte-identical. Additive
  and opt-in — the reader/writer API is unchanged.
- **`0.9.0`** — tables, data validation & conditional formatting, read AND write. `Worksheet.tables`
  / `dataValidations` / `conditionalFormatting` and the matching `SheetInput` fields carry the same
  records both ways, so they round-trip through `workbookToInput` (tables: name/range/columns/header
  & totals flags/style; data validation: all eight types + operators + prompt/error text + intuitive
  `showDropDown`; conditional formatting: highlight rules with an inline differential style, color
  scales, data bars, icon sets). The tolerant reader normalizes a foreign producer's out-of-spec
  table (odd name, mismatched column count, impossible totals row) into writer-legal shape so it
  re-saves instead of aborting. Hardened by a new property + mutation **fuzzing harness**
  (`@openjsxl/fuzz`, private) run over every reader/writer. Additive API — unused features emit
  nothing, so `.xlsx` output for existing input stays byte-identical.
- **`0.9.1`** — hostile-input & fidelity hardening on top of 0.9.0, no API change. **ST_Xstring
  escaping** so string content with XML-illegal characters (control chars, lone surrogates, a literal
  `_xHHHH_`) round-trips through Excel/openpyxl instead of corrupting or rejecting; **Unicode table
  names** read verbatim (non-English Excel locales' default names no longer mangled); a `MAX_CF_FORMULAS`
  shared bound. Plus **hostile-input hardening**: the streaming tokenizer is O(n) on any chunking, and
  every read is guarded against decompression bombs **by default** (a 2 GiB per-part ceiling + a 300×
  compression-ratio cap, tunable/disable-able via `ReadOptions.maxPartBytes` / `maxCompressionRatio`);
  a part or cell value that would exceed the runtime's max string length fails typed (`part-too-large`),
  never a bare crash. `.xlsx` output for existing input stays byte-identical.
- **`1.0.0`** — the stable API. Milestone M10 completes the workbook-level fidelity carry —
  **defined names**, sheet **autoFilter**, **protection** (sheet / workbook / per-cell locked+hidden),
  and **print setup** (margins, page setup, print options, header/footer) all read AND write and
  round-trip through `workbookToInput`; `Workbook.macroEnabled` flags an `.xlsm` whose VBA a rewrite
  drops. Then an **API freeze**: every public function reports failure with a typed `XlsxError` — the
  a1 helpers (`parseRef`, `columnToIndex`, …) included, so no bare `Error` reaches the public surface;
  the two entry points are collision-free (`openjsxl`'s `CellRef` vs `openjsxl/formula`'s `CellRefNode`),
  so a consumer can import both at once; every type named by a public signature is exported (guarded by
  an export-surface pin test); and `require()` resolves on Node ≥ 24 via a `"default"` export condition
  plus a `"./package.json"` subpath, with the ESM-only / Node ≥ 24 stance documented in the READMEs. The
  README becomes the **canonical documentation** — a full API reference for both entry points, verified
  against the built `.d.ts`. `.xlsx` output for existing input stays byte-identical (the freeze is
  additive / hygiene only). Follow semver from here.
- **`1.1.0` (in progress; not published)** — a minor release of correctness and hardening work on
  the 1.0 line. Input outside the cases below reads and writes as in 1.0.0.

  **New public members.**
  - `ReadOptions.verifyChecksums` (off by default) checks each fully read ZIP entry's CRC-32.
  - `ReadOptions.maxTotalBytes` (unlimited by default) bounds the decompressed bytes of every entry
    read through one archive; the read that would pass it fails with `part-too-large`.
  - `Workbook.date1904`, with an optional `date1904` argument on the `Workbook` constructor.
  - `WorkbookInput.date1904` and `StreamWorkbookInput.date1904`.
  - `EvalContext.date1904`, optional: the evaluator always sets it, and a hand-built context without
    it uses the 1900 system, so a 1.0.0 context literal compiles unchanged.

  **Changes that can break 1.0.0 code**, each with its migration.
  - *The writers read only own properties.* `writeXlsx`, `streamXlsx` and the bridge read only the
    own properties of cells, sheets, the workbook, the options object and table column entries. A
    field that a prototype supplies, such as a class getter or an `Object.create` record, fails with
    `XlsxError('invalid-input')` naming it; 1.0.0 wrote such fields. Table column entries must be
    plain objects: `null`, a string or a class instance now fails in both writers, where 1.0.0's
    buffered writer ignored the entry and took the name from the header row (its streaming writer
    already refused them). A key polluted onto `Object.prototype` stays ignored. A Proxy can still
    hide a field from the check, by serving it from its `get` trap alone or by reporting a plain
    prototype, and that field is not written. Migration: pass object literals that own their model
    fields.
  - *Read options are checked before a file is read.* Each limit (`maxPartBytes`,
    `maxCompressionRatio`, `maxTotalBytes`) must be a number of at least 0, `Infinity` for none, and
    `verifyChecksums` must be `true` or `false`; `undefined` and `null` keep the default. `NaN` (which
    lifted every per-entry cap in 1.0.0), strings (numeric ones included), negative numbers, `Number`
    objects and a non-boolean `verifyChecksums` now fail with `invalid-input` in `openXlsx`,
    `openXlsb`, `openOds`, `streamSheetRows` and `detectSpreadsheetFormat`. Detection now throws for a
    bad option, for any input, instead of returning a format or `undefined`.
    `EvaluateOptions.maxCellVisits` follows the same rule and fails with `FormulaError('unsupported')`
    before any sheet is read; on 1.0.0, `NaN`, `"abc"` or `{}` meant no budget. Migration: pass a
    number, `Infinity` or a boolean, or leave the option out; `Number(process.env.UNSET)` is `NaN`.
  - *A workbook's date-system key is read.* `WorkbookInput.date1904`, which 1.0.0 ignored, selects
    1904 for `true` and 1900 for `false` or `undefined`; any other value, `null` included, fails with
    `invalid-input`. Migration: pass a boolean or leave the key out.
  - *A numeric `t="d"` cell reads as empty.* A `t="d"` cell that holds a number (`45366`) reads as
    empty, and the bridge drops it; 1.0.0 read the number. `t="d"` holds ISO 8601 text, so conforming
    files need no migration.
  - *Escaped sheet names and Note authors are decoded.* A sheet name or legacy Note author is read
    as `ST_Xstring`, so `_x0041_` reads as `A`. When the decoded text holds a character the writer
    refuses, as `a_x0001_` does, the reader returns it, and `workbookToInput` followed by `writeXlsx`
    fails with `invalid-input` for the whole workbook; 1.0.0 returned the escape's text and rewrote
    the file. Migration: catch the error, or rename the sheet or author before writing.
  - *Pictures can share a buffer.* Pictures that use one image part share one `bytes` buffer,
    across drawings and sheets. Migration: treat `bytes` as read-only, and copy it before mutating
    or transferring it.
  - *Caller functions receive array constants as ranges.* A function registered through
    `EvaluateOptions.functions` receives an array-constant argument as a `RangeView` whose sheet is
    `""`, not as its first element. `parseFormula` throws a typed `FormulaError` parse error for a
    ragged array constant such as `{1,2;3}`. Migration: handle a `RangeView` argument as for a range.
  - *Inputs that now fail typed*, which 1.0.0 read, misread or wrote:
    - ZIP64 metadata, except an end record that only repeats the classic one, and multi-disk
      archives: `unsupported`, at open;
    - an inconsistent central directory, two entries that share bytes, or entry data that reaches
      the central directory or runs past the end of the file: `corrupt-zip`, at open, even when no
      read touches that entry;
    - an entry that ends short of its declared size: `corrupt-zip`. `detectSpreadsheetFormat`
      reports a ZIP file it cannot open as unrecognized (`undefined`);
    - CSV input with more than 16,000,000 non-empty cells within Excel's grid, or bytes that decode
      to more text than the runtime's longest string: `part-too-large`;
    - an `.xlsb` sheet with more than 16,000,000 cells, or more than 16,000,000 cells with a
      non-default style, or a workbook with more than 16,000,000 sheets, shared strings or cell
      formats: `part-too-large`;
    - an ODS document whose repeated cells and rows create more than 2,000,000 cells beyond their
      elements: `part-too-large`, where 1.0.0 stopped silently at 2,000,000 cells;
    - writing a number-format code longer than 255 characters (Excel's limit), as a cell style or a
      conditional-format style: `invalid-input`, in both writers.

    Migration: none for conforming files; these fail with `XlsxError`, as other malformed input does.
  - *Data the readers now leave out*, which 1.0.0 returned:
    - a later sheet whose worksheet part an earlier sheet, or a workbook-level part, already used,
      in `openXlsx` and `openXlsb`, with its scoped defined names. Its name stays free, so a later
      sheet with that name loads. A repeated comments or table part, a sheet-rels part read in
      another role, and a repeated shared-strings, styles or theme relationship read as absent;
    - in `.xlsx`, a later tab whose decoded name matches an earlier one ignoring case (`A` and `a`, or
      a raw `_x0041_` and `A`), with its scoped defined names; a formula naming it reads the kept tab.
      The comparison is `toLowerCase`, the rule `writeXlsx` uses to refuse two sheet names;
    - an `.xlsx` hyperlink with no resolvable external target and no usable location, a location
      that decodes to a character the writer refuses included;
    - a defined name whose `refersTo` holds an `_xHHHH_` escape that decodes to a character the
      writer refuses, such as `_x0001_` (tab, line feed and carriage return are allowed); 1.0.0
      returned the escape's text;
    - a number-format code longer than 255 characters: `numberFormat()` returns `undefined`, and the
      cell is not date-typed;
    - an ODS cell's comment (`office:annotation`) and its anchored shapes, which 1.0.0 read as the
      cell's text or hyperlink;
    - in `.xlsx`, cell values, styles, formulas, row properties and columns that sit inside a
      worksheet extension, under a foreign root or in a nested `sheetData`;
    - a picture listed twice: a drawing named twice by one sheet returns its pictures once, and a
      drawing named by several sheets gives its pictures to the first sheet in tab order only.

    Migration: none for files that follow OOXML and ODF.

  **Corrections.**
  - *ZIP container.*
    - Each fully read entry must reach its declared size. Short fails with `corrupt-zip`, and an
      overrun fails with `part-too-large`, as in 1.0.0; every inflate stops at the declared size,
      whatever the limits are. With `verifyChecksums`, a CRC-32 mismatch fails with `corrupt-zip`.
      A cancelled stream or an unread entry makes no size or checksum claim.
    - `openXlsx` and `openXlsb` read each worksheet, sheet-relationships, comments, table,
      shared-strings, styles and theme part at most once per open; `Worksheet.images()` reads each
      drawing and image part at most once per workbook.
    - `streamSheetRows` reads a stored (uncompressed) worksheet in bounded 16 KiB slices, where 1.0.0
      failed with a false `part-too-large`, and keeps a U+FEFF anywhere but at the document start.
    - `streamSheetRows` yields rows as it decodes them, so treat streamed rows as provisional until
      the iteration completes: a short entry or a checksum failure throws after the last row.
  - *Escaped text (`ST_Xstring`).*
    - Sheet names, legacy Note authors and defined-name `refersTo`; cell, data-validation,
      conditional-format and table formulas; and hyperlink locations are now encoded and decoded as
      `ST_Xstring` in both directions. A literal `_x0041_` is written `_x005F_x0041_` and survives
      Microsoft Excel, instead of becoming `A` or an external reference. A raw escape-shaped value
      that 1.0.0 wrote in these fields is read per OOXML, with no legacy mode. Fields without such
      text are written byte for byte as before. A Microsoft Excel for the web 16.0300 save-back
      fixture pins the corrected local reference.
    - A carriage return in a formula, a defined name, a table formula or a Note author survives a
      re-read, written as `_x000D_`; 1.0.0 wrote a raw CR byte, which XML parsers read as a line feed.
      A hyperlink location also carries a CR as `_x000D_`. openjsxl 1.0.0 and openpyxl 3.1.5 show
      `_x000D_` literally in formula fields, hyperlink locations and Note authors; SheetJS decodes it
      in formulas and defined names, but not in Note authors or a hyperlink's location.
  - *`.xlsx` reader.*
    - `t="d"` cells holding ISO 8601 date text (openpyxl's `iso_dates=True`) read as dates in
      `cell()`, `rows()`, `streamSheetRows`, the evaluator and the bridge; 1.0.0 read them as empty.
      Other text, numeric time-zone offsets such as `+01:00` and bare times read as empty, as in
      1.0.0.
    - A workbook's `date1904` flag reaches `Workbook.date1904`, the evaluator and the bridge.
    - Cell values, styles, formulas, row properties and columns come only from the worksheet's own
      `<sheetData>` and `<cols>`. Extension payloads never change cell values, and a cell's
      `style()` and `numberFormat()` agree with `cell()`.
    - Resolving column styles no longer takes time in proportion to columns times cells on a
      hostile sheet.
    - Excel 4.0 macro sheets read and bridge as in 1.0.0, written back as regular worksheets whose
      macro semantics are dropped.
  - *Writers.*
    - Each writer reads a sheet's `rows`, each row's length, each cell's style and each table
      header value once, and writes the values it validated, so a getter that changes cannot slip
      past validation.
    - A `WriteOptions.date1904` that holds a value overrides `WorkbookInput.date1904`. Only `true`
      selects 1904, as in 1.0.0, so `"false"`, `"0"`, `1` or `"yes"` keep the 1900 system; `null`
      and `undefined` defer to the workbook flag.
    - Sparse rows cost the cells written, not the row's length, and conditional-formatting blocks
      are written in linear time. A block list that grows during the write no longer adds an empty
      `<conditionalFormatting sqref="">`. Output bytes for valid input are unchanged.
  - *`.xlsb`.*
    - `BrtWbProp`'s 1904 flag sets the date system, so a 1904 workbook's dates no longer shift by
      1,462 days.
    - An all-zero dimension record, which 1.0.0 read as `A1`, resolves from the sheet's cell
      records and default row and column formatting: an empty sheet has no dimension (`undefined`;
      `.xlsx` still declares `A1`), and a sheet with cells gets its used range, such as `B4:D7`.
      Width-only or hidden-only default columns do not widen it, and an all-max dimension, which
      1.0.0 read as `XFD1048576`, is absent.
    - Short cell records with implicit columns, which producers such as SheetJS 0.20.3 write, are
      decoded; 1.0.0 returned only the first cell of each run.
    - A cell record cut before the end of its value reads as empty, with its style, instead of an
      invented value such as the first shared string or `true`.
    - The default style 0 is not stored per cell, so sheets of blank style-0 records read quickly
      in little memory.
  - *ODS.*
    - A negative duration keeps its sign: `-PT1H` reads as −1/24.
    - Cell text is read in full, where 1.0.0 cut it at 32,768 characters; only the extra spaces of a
      `text:s` run are capped, at 32,768 per cell. A stray row close, or a row nested inside a cell,
      no longer copies the previous row into the next.
  - *CSV.* `openCsv` reads large input in memory near its size, and input that crashed now reads.
    Every input that read before parses to the same cells.
  - *Formulas (`openjsxl/formula`).*
    - Shared formulas: dependents keep structured references verbatim; lowercase references shift
      and are written in canonical uppercase (a `$a$1` master's dependents read `$A$1`, where 1.0.0
      returned `$a$1`); a shift off the grid reads `#REF!`; an unbalanced `[` leaves the rest of the
      formula untranslated. Unicode names and full 3-D qualifiers stay unchanged, and a range whose
      second endpoint is sheet-qualified shifts the cells `parseFormula` reads.
    - The workbook's date system drives every serial and date conversion, `NOW` and `TODAY`
      included, so a 1904 workbook no longer evaluates 1,462 days off.
    - `COUNTIFS`, `SUMIFS` and `AVERAGEIFS` give `#VALUE!` for ranges of different shapes and for an
      incomplete criteria pair, and count every blank position a criterion matches.
    - Wildcard and relational text criteria match text only, and a criterion that refers to an
      empty cell, or an omitted one, means 0.
    - `SUMIF` and `AVERAGEIF` give what `SUMIFS` and `AVERAGEIFS` give on the same single-criterion
      data: a matched error at a blank criteria cell propagates (`SUMIF(A1:A2,"",B1:B2)` with A2
      blank and B2 `#N/A` is `#N/A`), the first error in row order wins, an excluded cell is never
      evaluated, and a single-cell sum such as `B1` is walked like a range.
    - Functions stored with Excel's `_xlfn.` or `_xlfn._xlws.` prefix evaluate like their bare
      names, and a caller function registered under the bare name answers the prefixed call.
    - A formula whose result is an empty cell evaluates to 0, as Excel shows it (`=A1`,
      `=IF(TRUE,A1)`, a defined name), so `ISBLANK`, `COUNT`, `COUNTBLANK` and `&` agree with Excel.
      `IFERROR` and `IFNA` return `""` for an empty value or fallback, and `INDEX` returns an empty
      target as empty inside a formula.
    - A cell that carries only a style is ignored: `COUNTA` does not count it, and approximate
      lookups and `MATCH` of 0 or `""` skip it.
    - An array constant passed to a function is a full array: `SUM({1,2,3})` is 6 and
      `SUMPRODUCT({1,2},{3,4})` is 11. A one-value parameter, an operator, and a cell whose formula
      ends in an array still use the first element.
    - A sheet name in a formula, and `evaluateCell`'s sheet, match ignoring case. `Data!Rate` uses
      the `Rate` scoped to Data, then the workbook-wide one, and a qualifier that names no sheet
      gives `#REF!`.
    - `&`, `SUBSTITUTE`, `REPLACE`, `UPPER` and `LOWER` give `#VALUE!` for text longer than 32,767
      characters, as `CONCAT`, `TEXTJOIN` and `REPT` already did, and a result too long for
      JavaScript gives `#VALUE!` instead of a bare `RangeError`. Stored text longer than the limit
      still passes through unchanged.
    - `TEXTJOIN` with ignore_empty `FALSE` keeps an empty item for every empty cell of a range
      (`"a,,c"`); over a large range with a delimiter it can now give `#VALUE!`.
    - Ranges find their cells through a sorted index, so formulas that each read a small range down
      a long sheet no longer slow quadratically. A file that lists rows or cells out of order is read
      in row order.
    - A criterion, or a text lookup value, longer than 255 characters gives `#VALUE!` in the
      `COUNTIF`, `SUMIF` and `AVERAGEIF` families, `VLOOKUP`, `HLOOKUP` and `MATCH`; wildcard matching
      takes time in proportion to the text.
    - A built-in function reads a single-cell reference argument as a reference, as the same cell
      written `A1:A1` already read: `SUM(D1,D2)` with D1 `"n/a"` and D2 2 is 2, not `#VALUE!`.
      `INDEX`, `VLOOKUP`, `HLOOKUP`, `SUMPRODUCT`, the lazy `IF`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`
      and `CHOOSE`, and caller functions still receive the cell's value.
  - *Documentation.* The README states the date contract (UTC wall-clock dates, serials below 61 one
    day off from Excel, the ranges that match Excel, times and durations), the read and write limits,
    and which readers show which escaped text. SECURITY.md describes the guards for untrusted files.
  - *Costs.* Reading `.xlsx` cells costs about 13–27% more than in 1.0.0, from the container
    scoping, and `formula()` and the geometry accessors scan more of the sheet. Writing styled cells
    costs about 30–36% more, from the own-property checks; numeric writes are within 1%.
  - *Hardening.* The seeded property and mutation fuzz harness now also runs the readers with
    checksums on and compares streamed with buffered rows; new tests pin the code added since
    1.0.0; CI keeps its code checks running after an audit failure and fails below a minimum test
    count; dependencies were refreshed and frozen before release.

  **Known limitations**, planned for a later release.
  - The read and styled-write costs above are not yet recovered.
  - The writer copies each part once more than it needs to before compressing it.
  - Workbook and worksheet attributes parse booleans and integers through two copies of the same
    rules; they agree today.
  - A legacy Note `authorId` written as `1e0`, signed or padded still selects an author.
  - A decoded Note author or sheet name that the writer cannot write (a control character, a lone
    surrogate, a `:` in a sheet name) is returned as is, and the bridge then fails typed for the whole
    workbook. The same holds for a hyperlink tooltip, display text or target, a sheet-protection
    password or a table style name that holds such a character.
  - ZIP entry names are compared case-sensitively for duplicates, while OPC compares part names
    ignoring case, and a relationship target in the wrong case does not resolve. A plaintext entry
    flagged as encrypted opens, and a local header whose name or method differs from the central
    one is not refused.
  - The writers check a sheet's row count once but re-read it while writing, so a getter that
    appends a row gets it written, past Excel's grid from the last row, and a Proxy whose `length`
    changes defeats the check. Streaming reads a row source's iterator method twice, and a forged or
    proxied `Date` throws a bare `TypeError`.
  - The reader applies no XML line-end or attribute normalization: a literal CRLF reads as `\r\n`,
    where Excel and openpyxl see `\n`, and a bridge then writes the CR.
  - CI does not exercise the built packages, `require(esm)`, the examples or an installed tarball,
    and runs on one Node version and one OS.
  - `writeXlsx` throws a bare `RangeError` for a sheet too large for one string, which a crafted
    xlsx of about 105 KB reaches through `workbookToInput`; `streamXlsx` checks the ZIP size limits
    only after compressing all of its input.
  - In a sheet whose rows omit `r`, formulas and the evaluator take a row's position from the row
    count instead of its cells' own references.
  - Comment text includes phonetic readings; Thai built-in date formats (ids 71–81) read as
    numbers; a theme can make `resolveColor` throw a bare `TypeError`; and a duplicated cell
    reference mixes one cell's value with another's style and formula.
  - A declared but missing sharedStrings part reads every `t="s"` cell as empty, and a missing styles
    part reads dates as serial numbers. A missing worksheet part drops its sheet, and missing
    comments, table and image parts are skipped, all without an error.
  - `<v>` text is parsed loosely (`"true"` reads false, `"0x10"` reads 16), and `openXlsx` on
    `.xlsb` bytes, or `openXlsb` on `.xlsx` bytes, returns an empty workbook.
  - UTF-16 XML parts, and CSV bytes that are not UTF-8, are decoded as UTF-8 without an error.
  - The writers ignore unknown keys (`merge` for `merges`, `{ date194: true }`), take
    `rowProperties` keys through `Number()`, and accept any `CsvReadOptions`.
  - Conditional-format integers and enums, some defined and table names (`a+b`, `R1C1`), sheet names
    with a leading or trailing apostrophe or named `History`, and cell text over 32,767 characters
    are written without the bounds Excel applies. A defined name that looks like a cell reference,
    such as `YTD2023`, is refused on write and dropped on read, so a bridged formula that uses it has
    no definition.
  - The writer's merge check takes about 20 s per million merges; one malformed or overlapping merge
    read from a file fails the whole bridge; and overlapping columns, two comments on one cell, and
    merges or an autoFilter that cross a table are written.
  - 1900-system serials below 61 are one day off from Excel, and dates outside Excel's range are
    written without an error.
  - A rewrite drops the theme part when only table styles use it, which can change table colors.
  - The writer does not write the switch that turns fit-to-page on, so a written or bridged sheet keeps
    `fitToWidth` and `fitToHeight` but has fit-to-page off.
  - `.xlsb` hyperlinks keep only external targets, an unresolved one breaks the bridge, a damaged
    row header or an off-grid cell record can overwrite valid cells, and an out-of-range
    shared-string index reads as `""`, where `.xlsx` reads empty.
  - `openXlsb` and `openOds` keep sheets whose names match ignoring case, so a later one is
    unreachable by name; ODS renames an empty sheet name, while XLSB keeps it.
  - CSV inference drops a leading `+`, rounds a decimal with more digits than a double holds, and
    reads `1e-400` as 0.
  - ODS hidden sheets read as visible, malformed dates and numbers are changed rather than dropped
    (`2024-02-30` reads as 2024-03-01), in-document links read as external targets, and `P1D`
    durations without a time part are dropped.
  - Evaluation nesting past 256 levels, such as a 300-term `1+1+…` sum or about 130 rows of
    `=IFERROR(A2,0)+1`, throws `depth-exceeded` and aborts `evaluateWorkbook` for every cell.
  - Number-like text (`"$4.00"`, `"50%"`, `"1,234"`, dates as text) is not coerced in arithmetic,
    and date criteria such as `">=1/1/2024"` compare as text. Formatted numeric criteria such as
    `">1,000"`, `">$1"` and `">=5%"` never match number cells, where Excel compares them as numbers.
  - These criteria rules may differ from Excel and are not verified against it: `"="` acts as
    `""`; `"<>"` excludes formula `""` and error cells; `COUNTIF(rng,"#N/A")` gives 0; numeric
    criteria do not match numeric text; `"<>0"` and `"<>TRUE"` match only cells of the same type;
    error cells match no criterion, `"<>*"` included; and `AVERAGEIFS` leaves blank criteria cells
    unmatched, where Microsoft's page treats them as 0.
  - `SUMIF` and `AVERAGEIF` do not extend a sum range smaller than the criteria range.
  - Off-grid references (`A0`, `ZZZ1`) read as empty, and a chain of more than 64 single-cell
    ranges reports `#CYCLE!`.
  - In a cycle, only the cell that closes it gets `#CYCLE!`, so an `IFERROR` around a cycle can
    return its fallback and results can depend on evaluation order.
  - Non-finite results (`SUM(1E308,1E308)` is `Infinity`) and forged error codes pass the evaluator
    boundary; injected `now` and `random` values are not checked, and `now: new Date(0)`, a value
    rather than a function, throws a bare `TypeError`; a caller function can change the shared
    context for later cells.
  - Arithmetic has no near-zero compensation (`1.333+1.225-1.333-1.225` is about -2.2e-16, where
    Excel shows 0), and `ROUNDUP`, `ROUNDDOWN`, `FLOOR` and `CEILING` act on binary noise.
  - `PROPER` treats only ASCII as letters, `MATCH` on a 2-D range returns a row offset, `TRUE()` and
    `FALSE()` give `#NAME?`, `CHAR` and `CODE` use Unicode, and some edge cases (`0^0`, `LOG(x,1)`,
    `FLOOR(x,0)`, `FIND("","abc",4)`, `1E+20&""`) differ from Excel.
  - `IFERROR(1/0,)` gives `""`, where `IF(FALSE,1,)` gives 0; `VLOOKUP` and `HLOOKUP` turn an empty
    target into 0 inside a formula; a multi-cell range formula such as `=G1:G3` shows its first value
    but hands the whole range to dependents; and `INDEX(A1,0)` gives the cell where
    `INDEX(A1:A1,0)` gives `#REF!`.
  - A shared-formula range with an off-grid endpoint reads as `B1:#REF!` rather than `#REF!`; the
    value is `#REF!` either way.
  - Macro, chart and dialog sheets read as worksheets with no marker of their kind.
  - In a shared formula, a column or row pair before `#SPILL!` or `#CALC!` keeps its text, and a
    dependent keeps a cell unshifted when the unquoted run holding it opens with no cell
    (`S!A1:A2:T!B1`, `T[x]:A1:S!B2`, `@A1:S!B1`, `MyName:A1:S!B2`).
  - Merges and hyperlinks inside a worksheet extension (`<extLst>`) are still read and bridged, an
    external hyperlink target included.
  - The bridge writes a `t="d"` date as a serial under the cell's own format, so a date under a
    non-date format reads back as a number and an unstyled one gains `mm-dd-yy`; bare `t="d"` times
    read as empty.
  - A row or cell written inside an open cell value becomes a real row or cell.
  - `.xlsb` limits its cell budget per sheet, so several sheets under the budget can still exhaust
    memory; rewriting a workbook takes time that grows with a font name's length times the rules or
    styled cells that use it.
  - Cell-level writer errors omit the sheet name; CSV drops rows and columns past Excel's grid
    without an error, and CSV and ODS spell a one-cell dimension `A1:A1`; `openjsxl/formula`'s types
    need `moduleResolution` `node16`, `nodenext` or `bundler`; and the `Workbook` and `RangeView`
    constructors appear in the public types.
  - Some tests still assert wall-clock bounds, and no test checks the evaluator against the cached
    values in the fixture corpus.

## Notes

- **Provenance.** To publish with npm provenance from CI later, run `npm publish --provenance`
  from a trusted GitHub Actions workflow; not required for a manual release.
- **Deprecating a bad release.** npm forbids re-publishing a version. If a release is broken,
  publish a patch and `npm deprecate openjsxl@<version> "use <newer>"`.
