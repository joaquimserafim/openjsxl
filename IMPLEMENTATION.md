# openjsxl — Implementation Plan

The feature-by-feature breakdown that backs [ROADMAP.md](./ROADMAP.md). Every unit of work
is a **feature** with a stable ID, context, scope, design notes, scoped **tasks**, and
**acceptance criteria**. This is the tracker — check tasks off as they land.

**Status legend:** ☐ todo · ◐ in progress · ☑ done
**Task format:** `- [ ]` open · `- [x]` done

> Rule of the project: *nothing gets implemented before it is defined, scoped, and broken
> into tasks here.* Detail is deep for the near-term milestones (M0–M3) and intentionally
> lighter for later ones — those are expanded when we reach them.

---

## Architecture (context)

### Packages

| Package | Public? | Responsibility |
| --- | --- | --- |
| `@openjsxl/core` | yes | The engine. Internally layered: `zip` → `xml` → `ooxml` → `reader` (and later `writer`). Zero runtime deps. |
| `openjsxl` | yes | Thin facade. Re-exports the core public API; future home of the `openjsxl/write` and `openjsxl/native` entry points. Bundles core so users install one package. |
| `@openjsxl/fixtures` | no (private) | Test corpus + a zero-dependency generator that emits real `.xlsx` files. |

### The read pipeline

```
.xlsx (bytes)
  → zip:   locate End-Of-Central-Directory, walk central directory, inflate parts on demand
  → ooxml: read [Content_Types].xml + _rels/.rels → workbook.xml → workbook.xml.rels
           (resolve sheets through the RELATIONSHIP GRAPH, never by filename)
  → ooxml: parse sharedStrings.xml once into an index→string table; read the date1904 flag
  → xml:   SAX-stream worksheets/sheetN.xml as events (never build a DOM)
  → ooxml: type each <c> via its t attribute + style index → discriminated-union Cell
  → reader: expose Workbook → Worksheet → cell('A1') / rows()
```

### Layering & the backend-swap interface

The hot path is `zip` + `xml`. Both sit behind narrow interfaces so a future
`@openjsxl/native` (napi-rs binding to the Rust `calamine` crate) or a WASM build can
implement the same contract and be selected at runtime via `optionalDependencies`, with the
pure-TS path as the universal fallback. **The native lane is never the only path** — that
would forfeit the browser/Deno/edge story, which is the differentiator.

### Conventions

- **Imports:** extensionless (`moduleResolution: Bundler`); `tsup` emits runnable ESM.
- **No runtime deps.** Decompression via `DecompressionStream('deflate-raw')`; strings via
  `TextEncoder`/`TextDecoder`.
- **Async-first.** Public entry points return Promises.
- **Testing:** Vitest; tests live in `__tests__/` folders beside the code they cover,
  run against TS source (no build in dev). Every format rule gets a fixture-backed test.
- **Errors:** throw typed errors with the part/ref that failed; never silently mis-parse.
- **Quality gate:** `biome check` + `tsc` typecheck + `vitest run` all green before a
  feature is marked done.

---

## M0 — Foundations

Goal: a skeleton that installs, typechecks, lints, and tests green, plus the pure
primitives that have no xlsx-format logic.

### F0.1 — Package skeleton ☑
**Context.** The repo currently has root tooling and `@openjsxl/core`'s manifest only. We
need the full three-package layout and the core source tree so later features have a home.
**Scope (in):** directory layout; `src/index.ts` + `types.ts`; per-package `tsconfig.json`
+ `tsup.config.ts`; the `openjsxl` facade; the `@openjsxl/fixtures` package shell.
**Scope (out):** any parsing logic (those are their own features).
**Tasks**
- [x] Create `packages/core/src/{zip,xml,ooxml,reader}/` with `index.ts` barrels.
- [x] `packages/core/src/types.ts` — the public `Cell` discriminated union, `CellType`,
      `SheetInfo`.
- [x] `packages/core/src/index.ts` — public surface (re-exports only what's public).
- [x] `packages/core/tsconfig.json` (extends base, `noEmit` for typecheck) + `tsup.config.ts`.
- [x] `packages/openjsxl/` — manifest (`dependencies: @openjsxl/core`), `tsconfig`,
      `tsup.config.ts` (`noExternal: ['@openjsxl/core']`), `src/index.ts` (`export * from
      '@openjsxl/core'`).
- [x] `packages/fixtures/` — private manifest, `src/index.ts` (`fixturePath`,
      `loadFixture`).
- [x] `pnpm install` clean; `pnpm check && pnpm typecheck && pnpm test` green (empty/`todo`).
**Acceptance.** Fresh clone → `pnpm install` → all three quality commands pass. ✅ met.

### F0.2 — Test fixtures & generator ☑
**Context.** A reader is only as trustworthy as its corpus. We need both *programmatic*
fixtures (deterministic, committed) and *real-producer* fixtures (Excel, LibreOffice,
Google Sheets) to catch real quirks.
**Scope (in):** a zero-dep generator that writes a valid `.xlsx` using **stored** (no
compression) zip entries + CRC32; one `basic.xlsx` covering string, number, date, boolean,
and a cached formula; a README documenting how to add real-producer files.
**Design notes.** Stored zip entries need no deflate at all — only correct local/central
headers, CRC32, and an End-Of-Central-Directory record. `[Content_Types].xml` should be the
first entry. Validate output with `unzip -l` and by opening in a real spreadsheet app.
**Tasks**
- [x] `packages/fixtures/scripts/generate.mjs` — CRC32 + stored-zip writer + OOXML parts.
- [x] Emit `data/basic.xlsx`: `A1` string, `B1` number, `C1` date (styled), `D1` bool,
      `A2` string, `B2` float (plus `E1` cached formula for the M1 reader).
- [x] Wire the root `pnpm fixtures` script; add `allowJs`/`checkJs:false` to root tsconfig
      so `scripts/**/*.mjs` is allowed by typecheck.
- [x] `data/README.md` — corpus policy + checklist for adding real-producer files.
- [x] Verify via `unzip -t` (CRC integrity, all 7 parts OK) + `unzip -l`. Opening in
      Excel/LibreOffice is a recommended manual spot-check (no GUI in CI).
- [x] `src/fixtures.test.ts` — CI guard asserting the committed fixture is a valid ZIP.
**Acceptance.** `pnpm fixtures` writes a deterministic `basic.xlsx`; `unzip -t` passes (valid
zip + correct CRCs). ✅ met (real-app open not verified in this environment).

### F0.3 — Primitive: A1 addressing ☑
**Context.** Cells are addressed in A1 notation; columns are **bijective base-26**
(A=1 … Z=26, AA=27 — no zero digit). Needed by the reader and every test.
**Scope.** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef` (+ `CellRef` type).
**Design notes.** Case-insensitive column letters; reject malformed refs; Excel's max column
is `XFD` (16384), max row 1048576.
**Tasks**
- [x] `ooxml/a1.ts` with the four functions and validation.
- [x] `ooxml/a1.test.ts`: boundary cases (`A`,`Z`,`AA`,`XFD`), 1..20000 round-trip, rejects.
**Acceptance.** Round-trips every index 1..20000; parses/format the corner refs. ✅ met.

### F0.4 — Primitive: date serials & epochs ☑
**Context.** Excel stores dates/times as serial numbers under one of two epoch systems.
Whether a number *is* a date is decided later by the style (F2.1), not here.
**Scope.** `serialToDate(serial, date1904?)` → `Date` (UTC-based).
**Design notes.** 1900 system anchors at `1899-12-30` to absorb Excel's phantom
`1900-02-29` leap-year bug (correct for all dates ≥ 1900-03-01 — the universal convention).
1904 system anchors at `1904-01-01`. Fractional serials encode time-of-day.
**Tasks**
- [x] `ooxml/dates.ts`.
- [x] `ooxml/dates.test.ts`: Unix-epoch serial 25569; modern date 43831 = 2020-01-01;
      1904 serial 42369 = 2020-01-01; fractional → time.
**Acceptance.** Matches known Excel serials across both epoch systems. ✅ met.

---

## M1 — Reader MVP (v0.1)

Goal: `await openXlsx(bytes)` → read a correctly-typed value from `A1` of a real Excel file.
This is the demo that sells the project.

### F1.1 — XML entity decoding ☑
**Context.** OOXML text uses the five predefined entities plus numeric character
references. The tokenizer (F1.2) decodes both element text and attribute values with this.
**Scope.** `decodeXmlEntities(input)`.
**Design notes.** Handle `&amp; &lt; &gt; &quot; &apos;` and `&#nn;` / `&#xHH;`; leave
unknown entities intact; guard against out-of-range code points.
**Tasks**
- [x] `xml/entities.ts` + `xml/entities.test.ts` (predefined, decimal, hex, unknown, no-`&`
      fast path).
**Acceptance.** Decodes all valid forms; passes through invalid ones unchanged. ✅ met.

### F1.2 — Streaming XML tokenizer ☑
**Context.** Worksheets can be huge; a DOM is the reason pure-JS incumbents are slow and
memory-heavy. We stream the OOXML subset as events.
**Scope (in):** `tokenize(xml)` → iterator of `{open,name,attrs,selfClosing}` /
`{text,value}` / `{close,name}`. **Scope (out):** DTDs, PIs beyond `<?xml?>`, namespace
resolution (prefixes like `r:id` matched literally), validation.
**Design notes.** Must handle self-closing tags, single/double-quoted attributes,
`xml:space="preserve"`, and entity decoding (F1.1) in text and attribute values. Operate on
a string first (decode bytes once with `TextDecoder`); a byte-level scanner is a later
optimization. Avoid per-token allocation in the hot loop where practical.
**Tasks**
- [x] `xml/tokenizer.ts` scanner (tags, attrs, text, self-closing, comments skipped).
- [x] Entity decoding wired for text + attribute values.
- [x] `xml/tokenizer.test.ts`: nested elements, attrs, self-closing, preserved whitespace,
      CDATA-free OOXML samples, malformed-input behavior.
- [x] Adversarial review (3 lenses) — fixed 3 real bugs: surrogate code-point decode
      (`entities`), `/` self-close swallowing markup, and unescaped `<` phantom element;
      plus BOM stripping. Each pinned with a regression test.
**Acceptance.** Tokenizes the `sheetN.xml` and `sharedStrings.xml` fixtures correctly. ✅ met
(realistic worksheet-row + shared-strings shapes tested inline; full extraction lands with
F1.3).

### F1.3 — ZIP / OPC container reader ☑
**Context.** An `.xlsx` is a ZIP (Open Packaging Conventions). We need entry lookup + on
-demand inflate, with **zero deps** via the platform.
**Scope (in):** `inflateRaw(bytes)` (DecompressionStream); `openZip(bytes)` → entry map +
`read(name)`. **Scope (out):** zip64, encryption, writing (F3.1).
**Design notes.** Parse the End-Of-Central-Directory record, then the central directory for
each entry's name / method / sizes / local-header offset. Method 0 = stored (verbatim),
method 8 = raw deflate (→ `inflateRaw`). Read lazily — only inflate parts the caller asks
for. Tolerate a trailing EOCD comment.
**Tasks**
- [x] `zip/inflate.ts` — `inflateRaw` via `DecompressionStream`, with an output-size cap; round-trip test.
- [x] `zip/central-directory.ts`: EOCD scan, central-directory walk, local-header parse.
- [x] `read(name)`: dispatch stored vs deflate (output bounded by the declared size).
- [x] `zip/__tests__/*` against `basic.xlsx` + hand-built deflate / data-descriptor / extra-field zips.
- [x] Adversarial review (3 lenses): confirmed correct on the real-world data-descriptor
      (GP-bit-3) layout; added a decompression-bomb cap, wrapped inflate errors with the part
      name, and explicit ZIP64 rejection. Deferred robustness items pushed to F2.4.
**Acceptance.** Lists every part of `basic.xlsx` and returns exact bytes for each. ✅ met.

### F1.4 — Relationship graph (rels) ☑
**Context.** **Relationships, not filenames, are the source of truth.** Sheets are found by
following `r:id` from `workbook.xml` into `workbook.xml.rels`.
**Scope.** `parseRels(xml)` → `Map<id, {id,type,target}>`; helpers to resolve a part's
relationships relative to its path.
**Design notes.** `[Content_Types].xml` → `_rels/.rels` (→ officeDocument =
`xl/workbook.xml`) → `xl/_rels/workbook.xml.rels`. Targets may be relative; resolve against
the source part's directory. `sheetId` is **not** a file mapping.
**Tasks**
- [x] `ooxml/rels.ts` — `parseRels` (via the tokenizer) + `resolveTarget` (relative,
      `..`/`.`, and package-absolute `/` targets); `Internal`/`External` target modes.
- [x] `ooxml/__tests__/rels.test.ts` — unit tests + a real-`basic.xlsx` walk.
**Acceptance.** Resolves `Sheet1`'s `r:id` to `xl/worksheets/sheet1.xml` via the graph. ✅ met.

### F1.5 — Shared strings table ☑
**Context.** Most text lives once in `xl/sharedStrings.xml`; cells of type `s` store a
zero-based index into it.
**Scope.** `parseSharedStrings(xml)` → `string[]`.
**Design notes.** Parse once. Handle plain `<si><t>` and rich-text `<si><r><t>…</t></r>`
(concatenate runs). Respect `xml:space="preserve"`. This is the first big consumer of the
streaming tokenizer. Phonetic guides (`<rPh>`, `<phoneticPr>`) carry an alternate reading,
not the value — their `<t>` text is excluded (matches openpyxl). Lifted `localName` into
`utils` since rels needs it too. Adversarial review hardened misnested input by holding the
OOXML invariant **one table entry per `<si>` start**: a stray/nested `<si>`/`<si/>`
finalizes the open item first, so it can't drop or shift a well-formed neighbour's index.
**Tasks**
- [x] `utils/xml-names.ts` — `localName` lifted from rels + `__tests__/xml-names.test.ts`.
- [x] `ooxml/shared-strings.ts` — plain + rich-text runs; phonetic exclusion; misnest-safe.
- [x] `ooxml/__tests__/shared-strings.test.ts` — units + real `basic.xlsx` + misnest cases.
**Acceptance.** Index → string matches the fixture's `sst` exactly, runs concatenated. ✅ met.

### F1.6 — Worksheet cell stream & typing ☑
**Context.** The core value: turn `<c>` elements into typed `Cell`s.
**Scope.** A worksheet streamer that yields rows of `Cell`; `decodeCell(raw, ctx)` mapping
the `t` attribute + content to the discriminated union.
**Design notes.** `t` values: absent/`n` = number (default), `s` = shared-string index,
`str` = cached formula string, `b` = boolean (`0`/`1` inside `<v>`), `e` = error,
`inlineStr` = text in `<is><t>`. Booleans/errors live **inside `<v>`** — don't read them as
numbers. Cells/rows are **sparse and may be unordered** — key by ref. Date detection is
deferred to F2.1 (until then a date-styled number reads as a number). The value channel is
gated by type (inline strings read `<is>`, everything else `<v>`) and marked present on
element-open, so an explicit empty `<v></v>`/`<is><t></t></is>` reads as `""` not blank, and
a stray cross-channel element can't pollute the value (adversarial-review fixes). Missing
`r` falls back to positional column/row per spec; misnested rows/cells finalize-on-reopen.
**Tasks**
- [x] `ooxml/cell.ts` `decodeCell` for all `t` variants (date typing stubbed to number).
- [x] `reader/worksheet.ts` `readRows` streamer over `sheetN.xml` using the tokenizer.
- [x] Tests: each `t` variant; sparse, out-of-order, positional, empty/`""`, channel-gating;
      real `basic.xlsx` integration.
**Acceptance.** Reads string/number/bool/error/inline/formula-cached values with correct
types from `basic.xlsx`. ✅ met.

### F1.7 — Reader public API ☑
**Context.** The ergonomic surface users actually touch.
**Scope.** `openXlsx(source)` → `Workbook`; `Workbook.sheets`, `Workbook.sheet(name)`;
`Worksheet.cell('A1')`, `Worksheet.rows()` async iterator. `source`: `Uint8Array | ArrayBuffer`.
**Design notes.** Async (decompression is async). Worksheet XML is decompressed eagerly so
`cell()` is sync, but parsed→indexed lazily on first `cell()`/`rows()` (a sheet you never
touch costs a decompression, not a parse). TS discriminated-union cells so
`cell.type === 'date'` narrows `cell.value`. `parseWorkbook` reads the `<sheet>` list
(name/r:id/visibility); the part is located through the relationship graph, never a guessed
filename. `sheet(name)` throws (lists available) for a clean happy-path type; an absent
`cell(ref)` reads `empty`. `Uint8Array` subarray views (byteOffset>0) are handled — `openZip`
is view-safe. True streaming stays deferred to F2.3.
**Tasks**
- [x] `ooxml/workbook.ts` `parseWorkbook` (sheet list) + `ooxml/__tests__/workbook.test.ts`.
- [x] `reader/workbook.ts` `openXlsx` wiring zip → rels → sharedStrings → worksheet; `Workbook`/`Worksheet`.
- [x] `cell('A1')` and `for await (const row of sheet.rows())`.
- [x] `reader/index.ts` + export from `core/src/index.ts` (flows through the `openjsxl` facade).
**Acceptance.** `(await openXlsx(bytes)).sheet('Sheet1').cell('A1').value` is the right
typed value. ✅ met.

### F1.8 — Vertical-slice tests ☑
**Context.** Lock the MVP behavior with end-to-end tests on a real file.
**Tasks**
- [x] Wire `@openjsxl/fixtures` into the facade tests (workspace devDep on `openjsxl`).
- [x] Slice through the **public** `openjsxl` entry: list sheets via rels; read
      string/number/bool; `.xlsx`→JSON records (date deferred to F2.1).
- [x] README quick-start: "`.xlsx` → JSON in < 50 LOC".
**Acceptance.** `pnpm test` exercises a real `.xlsx` end-to-end and is green. **Tag v0.1.** ✅
met — `packages/openjsxl/src/__tests__/slice.test.ts` drives the facade; `openjsxl` and
`@openjsxl/core` bumped to `0.1.0`.

---

## M2 — Reader hardening (v0.2)

Goal: correct dates, constant memory, and the common metadata real files carry.

### F2.1 — Styles & number-format date detection ☑
**Context.** Nothing on a cell says "date" — a number is a date iff its style applies a
date/time number format. This unlocks correct `date` cells.
**Scope.** `parseStyles(xml)` → `StyleTable.isDateStyle(styleIndex)`; integrate into
`decodeCell` so date-styled numbers become `Date` (via F0.4).
**Design notes.** `c@s` indexes `cellXfs` (NOT `cellStyleXfs`); each `xf` → `numFmtId`. IDs
< 164 are built-ins not listed in `<numFmts>` (date/time ones: 14–22, 27–36, 45–47, 50–58);
IDs ≥ 164 are custom in `<numFmts>` — `isDateFormatCode` strips quoted literals, `\`-escapes,
and `[bracket]` sections, then looks for `d m y h s`. `cellXfs[0]` is the implicit default
(omitted `s` ⇒ `s=0`). Promotion applies to numeric cells only; the workbook `date1904` flag
(from `<workbookPr>`) selects the epoch. The style table + flag ride in `DecodeContext`,
shared across sheets; the `s` index travels on `RawCell`.
**Tasks**
- [x] `ooxml/styles.ts` (cellXfs + numFmts + built-in date ranges + format-code date sniff).
- [x] Integrate into `decodeCell`; thread `s` via the worksheet streamer; honor `date1904`.
- [x] Tests across built-in and custom date formats + a non-date number with `s` + 1904 +
      real `basic.xlsx`.
**Acceptance.** `C1` of `basic.xlsx` reads as a `Date`; a plain number with a style stays a
number. ✅ met.

### F2.2 — Constant-memory streaming reader ☑
**Context.** Worksheets and shared strings can be huge; peak memory must not scale with file
size.
**Scope.** Stream `sheetN.xml` from the zip without materializing the whole part; expose
row-at-a-time iteration end-to-end.
**Design notes.** Drive `DecompressionStream`'s readable directly into the tokenizer;
chunk-boundary-safe scanning. Kept **non-breaking**: sync `cell()` needs the part
materialized, and decompression is async-only, so streaming is a *separate* path
(`streamSheetRows`) rather than a change to `openXlsx`/`cell()`. The row state machine was
extracted into a push-based assembler shared by `readRows` (string) and `streamRows`
(chunks). `createXmlStream` buffers to a `safeBoundary` so no tag/comment/CDATA/PI/entity
straddles a cut; `TextDecoder({stream})` reassembles multi-byte UTF-8. Shared-strings
interning deferred (sharedStrings is read eagerly; usually far smaller than the worksheet).
**Tasks**
- [x] Streaming inflate (`inflateRawStream`) + zip `readStream`; chunk-fed `createXmlStream`.
- [x] `streamRows(chunks)` + public `streamSheetRows(source, sheetName?)`.
- [x] Memory test: read a 100k-row generated sheet streaming in fixed-size chunks.
**Acceptance.** Reads a 100k-row generated sheet with roughly constant memory. ✅ met
(`streamRows` holds ~one row; verified by the 100k-row + multi-byte + every-chunk-size tests).

### F2.3 — Common metadata ☑
**Scope.** Merged cells, hyperlinks, comments, per-cell number-format strings, sheet
visibility/dimensions.
**Tasks**
- [x] `Worksheet.mergedCells`; `Worksheet.hyperlinks` (resolved via worksheet rels);
      `Worksheet.comments`; `Worksheet.numberFormat(ref)` (chosen over a cell field);
      `SheetInfo.visible`/`Worksheet.dimension`. Each accessor adversarially reviewed.
**Acceptance.** Each feature verified against a real-producer fixture. ✅ met — merges,
hyperlinks, number formats (custom), visibility, and dimension are backed by committed real
files; **comments' real-producer check is local-only** (Apache-2.0 POI, CI-skipped by
design), covered in CI by inline units. Known limitations tracked for later: legacy-only
comments (no threaded), column/row default styles not resolved (#22), no formula text.

### F2.4 — Robustness ☑
**Scope.** Sparse/unordered cells, missing `<dimension>`, large/odd shared strings,
malformed input.
**Tasks**
- [x] Tolerate missing optional parts; clear typed errors on corrupt input; fuzz the
      tokenizer/zip parser. (`XlsxError`; missing-part tolerance; seeded fuzz over `tokenize`,
      `createXmlStream`, `openZip`, `openXlsx`, `streamSheetRows` asserting only `XlsxError`
      ever escapes — no bare `Error`/`TypeError`/`RangeError`, no hang.)
- [x] Zip hardening deferred from F1.3: commit a real Excel/LibreOffice/Sheets `.xlsx`
      fixture (closes the F1.3 review's S7); a configurable max part size; decide policy for
      duplicate entry names (rejected) and directory entries (`name/` placeholders, skipped).
- [x] Resolve column/row default styles (#22): a cell's effective format is `cell s` → row
      default (`<row s customFormat>`) → column default (`<col … style>`) → style 0, feeding
      both date detection and `numberFormat`.
**Acceptance.** No crashes on a corpus of malformed/edge-case files; errors are actionable.
The fuzz pass + an adversarial review found one bare-throw (an overflowing column ref reaching
`formatRef` as `Infinity`); fixed by rejecting the overflow in `columnToIndex` so the reader
falls back to positional addressing. A second review pass on #22 found the two style accessors
disagreed on no-`r` cells (positional in the assembler, skipped in `numberFormat`); fixed so both
address cells identically.

---

## M3 — Writer MVP (v0.3)

Goal: write a valid `.xlsx` Excel opens without "repair", round-tripping values and types.

### F3.1 — ZIP / OPC writer ☑
**Scope.** `writeZip(entries)` builds a ZIP from named byte-parts: correct local/central headers +
EOCD, CRC-32 per entry, `CompressionStream('deflate-raw')` with store-when-not-smaller fallback.
Internal only — the public `openjsxl/write` surface arrives in F3.2.
**Design notes.** Mirror image of `zip/central-directory.ts`; CRC-32 and byte layout follow the
generator's stored-zip code (F0.2). Deflate via the platform `CompressionStream` (symmetric with
the reader's `DecompressionStream`), so the writer stays zero-dep and runtime-agnostic — no
`node:zlib`, no hand-rolled inflate. Each part is deflated (method 8) and kept only when strictly
smaller than the raw bytes, else stored (method 0). Buffer each part in full so CRC + both sizes go
straight into the local header (no data descriptor, general-purpose bit 3 = 0). Fixed DOS timestamp
⇒ byte-identical output for identical input. Refuse anything that would need ZIP64 (`XlsxError`
`unsupported`) rather than emit a sentinel the reader misreads.
**Tasks**
- [x] `writer/crc32.ts` (table-driven CRC-32) + known-vector test.
- [x] `writer/deflate.ts` (`deflateRaw` via `CompressionStream`, mirrors `inflate.ts`).
- [x] `writer/zip.ts` (`writeZip`: store + deflate + headers + EOCD; ZIP64 + duplicate guards).
- [x] `writer/__tests__/` — round-trip `writeZip` → `openZip` (stored + deflate), determinism, guards.
**Acceptance.** Output re-reads byte-identically through F1.3 (`openZip`), both compression methods.

### F3.2 — Minimal workbook writer ☑
**Scope.** `writeXlsx(workbook, options?)` — public, flat/declarative API (symmetric with
`openXlsx`). Emits the minimal valid part set: `[Content_Types].xml`, `_rels/.rels`,
`xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/worksheets/sheetN.xml`, and `xl/styles.xml`
**only when a date is present**. Values + types (string/number/bool/date), multiple sheets.
**Design notes.** Input is plain data — sheets of row-major `CellValue[]` (`string | number |
boolean | Date | null | undefined`); the OOXML type is inferred from the JS value. **Inline
strings** (no shared-strings table to buffer). Escape `& < >` (and `"` in attrs);
`xml:space="preserve"` for edge whitespace; numbers via `String(n)` (shortest round-trippable
form; non-finite rejected); dates via `dateToSerial` + numFmtId 14 style (`s="1"`); no
`calcChain.xml`. Exported from the package index (tree-shakeable), not a subpath — `openjsxl/write`
can be added later additively. Validates inputs (≥1 sheet; sheet names 1–31 chars, no
`\ / ? * [ ] :`, unique case-insensitively) and throws `XlsxError('invalid-input')` rather than
emit a file Excel would "repair".
**Tasks**
- [x] `dateToSerial` (inverse of `serialToDate`) + `invalid-input` error code.
- [x] `writer/xml.ts` (escape/preserve), `writer/sheet.ts` (worksheet XML, dimension, typing).
- [x] `writer/workbook.ts` — `writeXlsx`, validation, part assembly; exported from core index.
- [x] Tests: write → `openXlsx` round-trip (types/values/sheets/sparse/1904/dimension), validation.
**Acceptance.** Round-trips through `openXlsx` **and** opens in openpyxl (independent reader) —
verified for strings (incl. `& < >`, edge whitespace), ints/floats/negatives, booleans, date +
datetime, sparse cells, multi-sheet. Real Excel/LibreOffice fidelity is F3.3.

### F3.3 — Round-trip fidelity ☑
**Scope.** `workbookToInput(workbook)` — the reader→writer bridge: turns an open `Workbook` into
`writeXlsx` input (each populated cell placed at its own A1 ref; sheets sparse via array holes),
enabling read → modify → write. Public, exported from the package index.
**Fidelity** is scoped to values, types, and sheet names/order (the writer's supported set).
Documented non-fidelity (not silently mangled): formulas keep only their cached value; error cells
become their text; merges/hyperlinks/comments/custom number formats/sheet visibility drop (M4).
**Tasks**
- [x] `writer/from-workbook.ts` (`workbookToInput`) + export.
- [x] Fidelity tests: read→write→read stability over `basic.xlsx`/`minimal.xlsx`; bridge over the
  supported value matrix (sparse, multi-sheet, dates); a golden XML-string pin of the wire format.
- [x] Reader robustness fix (found by review): a date-styled serial outside JS's Date range now
  reads as a `number`, not an Invalid Date (which would crash the writer on round trip).
- [x] Independent validation: `basic.xlsx` → `workbookToInput` → `writeXlsx` re-opens in openpyxl
  with identical values.
**Acceptance.** Round-trip lossless for the supported value set — verified against real fixtures and
openpyxl. **Release (separate step): tag v0.3** — bump `0.3.0`, README/PUBLISHING, writeXlsx example.

---

## M4 — Styles (v0.4)

Goal: read **and** write fonts, fills, borders, alignment, and number formats — plus sheet
geometry and the structural metadata the writer still drops. One shared `CellStyle` type is the
read model, the write model, and what the bridge carries, so styled round-trip stays a
pass-through. *The SheetJS-Pro killer.*

**Milestone-wide decisions** (design panel, unanimous): write input is a per-cell union
`CellInput = CellValue | { value, style? }` (backward compatible — every such object throws
`invalid-input` today); read API is a lazy `sheet.style(ref)` sibling of `numberFormat(ref)`
reusing the #22 effective-style precedence (cells stay flat — no styleId on `Cell`); colors are
the **raw** union `{rgb} | {theme, tint?} | {indexed} | {auto}` with **no theme1.xml parser**
(openpyxl's model; the only one that round-trips faithfully — resolving to rgb is lossy); number
formats travel as **code strings**, never ids. Deferred (named, not silent): comments write
(needs a legacy VML drawing part for Excel to display them → 0.5), theme parsing / resolved-rgb
helper, gradient fills, rich-text runs, cell protection, named-style (`cellStyleXfs`)
inheritance, conditional formatting, autofilter, outline levels, split (non-frozen) panes.

### F4.1 — Full style read model + `Worksheet.style(ref)` ☑
**Scope.** Extend `parseStyles`' single tokenize pass to also collect `<fonts>`, `<fills>`,
`<borders>`, and each cellXfs `<xf>`'s component ids + `<alignment>`. `StyleTable` gains
`cellStyle(styleIndex)` — materialized lazily, cached per xf index (reference-stable; this cache
is the interning unit the bridge reuses). New `Worksheet.style(ref): CellStyle | undefined`
resolved through the same lazy cell-style map as `numberFormat(ref)`, so the two can never
disagree. Hot value path (`isDateStyle`/`formatCode`) untouched.
**Design notes.** Component resolution ignores `apply*` flags and `cellStyleXfs`/`xfId`
inheritance (openpyxl-compatible). A fully-default xf resolves to `undefined` (the bridge then
keeps the cell a bare value). Colors kept verbatim as the raw union; gradient fills and exotic
underline/rotation values degrade to `undefined`, documented.
**Tasks**
- [x] `CellStyle`/`FontStyle`/`FillStyle`/`BorderStyle`/`Alignment`/`Color` types (shared).
- [x] `parseStyles`: fonts/fills/borders/xf components + `cellStyle(i)` with caching.
- [x] `Worksheet.style(ref)` accessor + export from core index.
- [x] Styled fixtures (real-producer `openpyxl-styled.xlsx`, provenance in data/README); tests for
  fonts/fills/borders/alignment/colors incl. rgb/theme+tint/indexed; reference-stability.
- [x] Review fixes: section closes flush dangling font/fill/border builders (misnested input no
  longer swallows cellXfs or grafts `<dxf>` children onto cell styles); font dispatch name-gated.
**Acceptance.** Styled fixtures read back verbatim; `style()` and `numberFormat()` agree on
inherited (row/col-default) styles; unstyled cells → `undefined`.

### F4.2 — Styled-cell write input + style interner ☑
**Scope.** Widen rows to `CellInput = CellValue | StyledCell` (`{ value: CellValue; style?:
CellStyle }`, `value` required but nullable — a styled blank emits `<c r s/>`, which
`worksheetXml` must stop pruning). Internal style registry interns styles structurally into
minimal deduped `cellXfs` and emits the full `styleSheet` (fonts/fills/borders/numFmts/cellXfs);
a static default `theme1.xml` part is emitted only when a theme color is written.
**Design notes.** Discrimination is total: null/undefined → empty; string/number/boolean/Date →
bare; any other object must be a `StyledCell` (missing `value`, unknown keys, bad enums, or
malformed colors → `invalid-input` naming the ref). Excel structural invariants: fills 0/1
reserved (none/gray125), empty font+border at index 0, `cellStyleXfs` + Normal cellStyle
present, solid fills colored via `fgColor`. **Hard gate: bare-value input reproduces the v0.3
output byte-for-byte** (golden pins), date and no-date cases.
**Tasks**
- [x] `CellInput`/`StyledCell` types; validation; `worksheetXml` styled + styled-blank cells.
- [x] Style registry (canonical structural key → xf index; deterministic emit order) + stylesheet
  emission; default `theme1.xml` (standard Office theme, from openpyxl output) when needed.
- [x] Tests: styled write → `style(ref)` deep-equals input; byte-identical golden for bare input
  (PLUS out-of-band full-archive byte-compare vs the pre-F4.2 dist: 5/5 identical); validation
  rejections; interning (identical styles share one xf); openpyxl reads all components back.
- [x] Review fixes (4 bugs, regression-pinned): `font.name` now isXmlSafe-gated; theme/indexed
  capped at u32 (no exponential notation in integer attrs); style components must be strictly
  plain objects (a Map's prototype `.size` getter had validated as a font size); validators
  single-read every property (a getter could pass checks then inject markup on the emission read).
**Acceptance.** Styled cells re-read exactly; v0.3 bytes unchanged for unstyled input; openpyxl
opens styled output with matching formatting.

### F4.3 — Number-format write (built-in reverse map + custom ids ≥ 164) ☑
**Scope.** Activate `CellStyle.numberFormat` in the writer as a format **code string**. Codes
exactly matching `BUILTIN_FORMATS` reverse-map to their id (no `<numFmts>` entry); others intern
from 164 up in deterministic first-encounter order. A `Date` with a user code keeps it (implicit
id 14 only when absent). Re-read typing flows through `isDateFormatCode` — a number written with
a date code re-reads as `date` (Excel-faithful; documented).
**Tasks**
- [x] `BUILTIN_CODE_TO_ID` reverse map; registry numFmt interning; `<numFmts>` emission.
- [x] Tests: built-in maps flat; custom round-trips verbatim via `numberFormat(ref)`; date
  interplay both directions; locale-id codes documented as non-representable.
- [x] Review fixes (2, regression-pinned): `isDateFormatCode` now strips quoted literals /
  escapes / skip-fill tokens BEFORE the elapsed-time sniff (a quoted `"[h]"` no longer date-flips
  a written number — pre-existing F2.1 heuristic bug F4.3 made reachable); `escapeAttr` emits
  tab/LF/CR as character references so attribute-value normalization can't silently rewrite
  format codes (or sheet/font names) for conforming readers.
**Acceptance.** `numberFormat(ref)` returns the written code verbatim after round-trip.

### F4.4 — Bridge carries styles (round-trip fidelity) ☑
**Scope.** `workbookToInput` attaches `sheet.style(cell.ref)` per populated cell — `{value,
style}` only when a style exists (unstyled v0.3-era workbooks produce identical input → the
byte-identical path), including styled *empty* cells (`<c s/>`) it currently drops. Signature
unchanged. README fidelity table rows move to "lossless".
**Design notes.** Documented flattening: row/column-default styles resolve into per-cell styles;
files authored under a **custom theme** keep `{theme,tint}` indices but re-render against our
default theme after rewrite — documented loudly. Property test: bridge output must always pass
`writeXlsx` validation.
**Tasks**
- [x] Bridge style attachment (incl. styled empties); acid test on the openpyxl-authored fixture
  (theme+tint, custom numFmts deep-equal across the bridge); README fidelity table update.
- [x] Corpus property test: EVERY readable fixture round-trips losslessly or fails TYPED — it
  found and pinned two pre-existing bugs (LibreOffice's explicitly-interned custom "General"
  format read asymmetrically; a bare-`Error` crash on unaddressable cell refs).
- [x] Review findings (5, all verified empirically after the verify agents rate-limited; fixed +
  pinned): the reader's style model now DEGRADES every value the writer would reject — shared
  bounds single-sourced in ooxml/styles (`HEX_COLOR` rgb gate, u32 caps on theme/indexed,
  `MAX_INDENT`, empty/XML-unsafe font names, empty format codes; `isXmlSafe` moved to utils) —
  so the bridge can never crash on a style the reader produced; plus Excel grid caps
  (`MAX_ROW`/`MAX_COL` in a1.ts) in the bridge AND the writer — a hostile ref like
  `A99999999999999` used to become `rows.length` and spin the writer for hours.
- [x] Second full adversarial pass over the fixed changeset (3 CONFIRMED / 0 refuted, fixed +
  pinned): custom format codes were the one field missing the `isXmlSafe` degrade (a control
  char via `&#1;` in `formatCode` still crashed the bridge); case-variant duplicate refs (`A1`
  vs `a1` — two reader identities, one grid slot) silently vanished a value — now a typed
  refusal, while same-spelling duplicates keep the reader's own last-wins semantics.
**Acceptance.** read → bridge → write → read gives deep-equal `style(ref)` for the supported
style set (verified per-cell over the whole fixture corpus + openpyxl cross-check of the
bridge-rewritten file); unstyled files rewrite byte-identically (even with dates).

### F4.5 — Sheet geometry: column widths, row heights, hidden, freeze panes ☑
**Scope.** Read: three lazy accessors in the `mergedCells` idiom — `columns`
(`{min,max,width?,hidden?}` from `<cols>`), `rowProperties` (`Map<row, {height?,hidden?}>` from
`<row>` attrs, dedicated scan off the hot path), `freeze` (`<pane state="frozen">`; split panes
read `undefined`). Write: matching `SheetInput.columns/rowProperties/freeze`; schema order
`sheetViews` → `cols` → `sheetData`; property-only rows emit cell-less `<row>`.
**Tasks**
- [x] Reader accessors + types (bounds shared with the writer, out-of-range degrades); writer
  emission + validation (schema order `dimension→sheetViews→cols→sheetData`; property-only rows
  emit cell-less `<row>`; no-op geometry normalizes to the exact geometry-free bytes); bridge
  carries geometry, and the corpus property test snapshots it for every fixture.
- [x] Tests: real-producer fixture (`openpyxl-geometry.xlsx`, provenance in data/README) reads
  exact widths/heights/hidden/freeze; write → re-read equal; openpyxl confirms frozen pane +
  widths on OUR output; byte-identity 5/5 held through the emission restructure.
- [x] Review fixes (4 CONFIRMED = 2 bugs + hardening, regression-pinned): `parseFreezePane` is
  now scoped to `<sheetViews>` (a saved Custom View's `<pane>` after `sheetData` fabricated a
  freeze the active view didn't have); `parseRowProperties` parses `r` with the row assembler's
  EXACT rule (`parseInt` + positional fallback — `Number()` divergence migrated heights to
  phantom rows on tolerated-malformed files); `parseColumnProps` stops at `<sheetData>`.
**Acceptance.** Geometry round-trips; Excel/LibreOffice show frozen header and sized columns.

### F4.6 — Structural metadata write: merges, hyperlinks, visibility ☑
**Scope.** `SheetInput` gains `merges` (A1 ranges → `<mergeCells>`; malformed/single-cell/
overlapping rejected — Excel repair-prompts on overlap), `hyperlinks` (reader-mirroring records →
`<hyperlinks>` + the writer's **first per-sheet rels part**, `TargetMode="External"` for
targets), and `state: 'visible'|'hidden'|'veryHidden'` (≥1 sheet must stay visible). Reader
`SheetInfo` gains `state` additively (`visible` boolean retained). Bridge carries all three.
**Tasks**
- [x] Merges + validation; hyperlinks + per-sheet rels wiring; visibility + guard;
  `SheetInfo.state` + `Worksheet.state`; bridge; tests (re-read via
  `mergedCells`/`hyperlinks`/`state`; openpyxl reads OUR merges/links/states warnings-as-errors
  and its own `openpyxl-metadata.xlsx` fixture reads back verbatim; byte-identity 5/5).
- [x] Merge-overlap detection is a row sweep, not O(n²): actives pruned past their last row all
  share the current first row, so non-overlapping actives are column-disjoint (≤16 384 alive) —
  a crafted million-merge file rejects in ms, not hours (the F4.4 grid-cap lesson applied).
- [x] `activeTab`: when the FIRST sheet is hidden, emit `<bookViews><workbookView activeTab=N/>`
  aimed at the first visible sheet (Excel's default 0 would open onto a hidden tab; openpyxl
  does the same) — all-visible workbooks emit no `bookViews` and keep exact pre-F4.6 bytes.
- [x] Review fixes (2 bugs, regression-pinned; verifier fleet rate-limited so each claim was
  re-verified empirically by hand): `sheet.state` was read up to 4× and interpolated unescaped —
  a value-flipping getter injected attributes into `workbook.xml` and dodged the all-hidden
  guard; `validate()` now reads state ONCE and returns the resolved array emission uses. The
  reader surfaced `target:""` from a crafted `Target=""` rel while the writer melted it away
  (silent round-trip change); `parseHyperlinks` now gates empty target like empty location.
**Acceptance.** The bridge's v0.3 drop-list shrinks to: comments, formulas, error cells.
**Release (separate tail commit): tag v0.4** — bump `0.4.0`, README fidelity table + styled
example, PUBLISHING note.

---

## M5 — Fidelity + streaming writer (v0.5)

**Theme.** Close the remaining `.xlsx` fidelity gaps the bridge still documents (comments,
formulas-as-text, custom themes) and ship the constant-memory writer — then make the "fast"
claim measurable. **Re-scope vs the original outline (owner-approved with this section):**
formula TEXT preservation is pulled forward from M8 (evaluation stays M8) and a benchmark
harness is pulled forward from 1.0, because both close competitive gaps documented in the
0.4 analysis; **images and the `@openjsxl/native` napi-rs/WASM lane move to M6** — images
need drawingML (a milestone of their own bundled with formats work), and the native lane
should follow, not precede, published pure-TS benchmarks.

**Standing gates for every M5 feature** (see CLAUDE.md for the full contract): biome by
exit code + tsc + vitest green; byte-identity for input not using the new feature;
the corpus property test extended to any newly-carried state; openpyxl cross-validation
both directions; adversarial review (finders + refuting verifiers) with confirmed findings
fixed and pinned before the commit is proposed.

**Dependency order for hand-off:** F5.2 → F5.3 → F5.4 build on the buffered writer's
existing emission helpers and are independent of each other after F5.2's per-sheet-rels
refactor; F5.1 (streaming) lands AFTER them so it reuses one final set of helpers; F5.5
(benchmarks) needs F5.1 for the streaming numbers. Suggested order: F5.2, F5.3, F5.4,
F5.1, F5.5 — one feature per session, each proceed-gated.

### F5.2 — Comments write (legacy VML) + bridge carry ☑
**Context.** The reader parses legacy comments (`xl/commentsN.xml`: authors + rich-text
runs concatenated to plain text) since F2.3; the bridge documents dropping them. Excel only
SHOWS a comment if the workbook also carries the legacy **VML drawing** part with a shape
per comment — a comments part alone reads back through openpyxl but renders nothing in
Excel (this is why openpyxl emits both).
**Scope (in).** `SheetInput.comments?: readonly Comment[]` (`{ref, author?, text}` — the
reader's exact shape). Emission per sheet with comments: `xl/commentsN.xml` (deduped,
ordered authors table + `<commentList>`), `xl/drawings/vmlDrawingN.vml` (one
`<v:shape>` per comment, hidden by default, anchored beside its cell — copy openpyxl's
anchor arithmetic), the worksheet element `<legacyDrawing r:id/>` (schema position: after
`hyperlinks`/`pageMargins` block — verify against CT_Worksheet sequence), rels entries for
both parts, and content types (`Default` extension `vml` + per-part comments `Override`).
**Scope (out).** Threaded (modern) comments — separate parts, deferred; documented. Rich
text runs — plain text only, matching the reader.
**Design decisions (made).**
- The per-sheet rels part is no longer hyperlinks-only: refactor `WorksheetResult` to
  return a list of rel entries (type + target + mode) so hyperlinks (F4.6), comments, and
  vmlDrawing allocate non-colliding `rId`s from one counter. This refactor is the reason
  F5.2 lands first in M5.
- Validation mirrors F4.6 hyperlinks: canonical single-cell `ref` (no ranges), single-read
  properties, unknown-key rejection, `isXmlSafe` on author/text, author optional.
- Author dedup: authors table is first-occurrence-ordered unique authors; a comment with
  no author points at a shared `""` author entry (matches openpyxl).
- VML is written from one minified template with per-shape substitution; byte-determinism
  as everywhere.
**Tasks**
- [x] Per-sheet rels refactor (`WorksheetResult.rels: SheetRel[]`, one rId counter; hyperlinks
      keep their exact bytes when no comments present — F4.6 tests unchanged, byte-identity 7/7).
- [x] comments + VML emission, validation, content types; `Worksheet` re-reads its own
      output verbatim (`comments` accessor).
- [x] Bridge carries `comments`; corpus snapshot gains `comments`; drop-list shrinks to
      formulas + error cells.
- [x] Fixture: `openpyxl-comments.xlsx` added (two resolved authors + an author-less comment;
      every commented cell valued so no empty-anchor placeholder); provenance in data/README.md.
- [x] openpyxl reads our comments (author + text) warnings-as-errors; adversarial review.
**Acceptance.** A written comment is visible in Excel/LibreOffice on hover (openpyxl
proxy-verified: both `comment.text` and `comment.author` match); round-trip lossless.
**Landed (uncommitted, awaiting owner approval).** `SheetInput.comments`; per-sheet
`xl/commentsN.xml` + `xl/drawings/vmlDrawingN.vml` (hidden note shapes, 0-based Row/Column, no
explicit `<x:Anchor>` — matches openpyxl) + `<legacyDrawing r:id>` after `</hyperlinks>`; the
per-sheet rels part is now generic so hyperlinks/comments/vmlDrawing share one non-colliding rId
counter. **Design note:** comments on an otherwise-empty cell keep the comment but drop the blank
anchor `<c>` (like every unstyled empty cell) — pinned by a test. Gate: biome 0 / tsc 0 /
**397 tests** / byte-identity 7/7 vs pre-F5.2 build / openpyxl both ways / all emitted parts (incl.
VML) well-formed. **Adversarial review: 0 findings** — 4 finders (spec-conformance, round-trip/
bridge, hostile-input, algorithm-correctness) with deep empirical probing surfaced nothing;
independently re-verified byte-identity, TOCTOU single-read, dup-ref tolerance (openpyxl accepts,
no repair), and multi-sheet VML idmap reuse (matches openpyxl).

### F5.3 — Theme fidelity: parse, resolve, carry ☑
**Context.** Colors are kept RAW (`{theme, tint?}` never resolved) — correct for
round-trip, but consumers can't get an RGB without their own theme parser, and the bridge
re-renders custom-theme files against the standard Office theme (documented flattening).
**Scope (in).** (a) `ooxml/theme.ts`: parse `theme1.xml`'s `<a:clrScheme>` into the
12-slot color table — **including the OOXML quirk that theme indexes 0/1 are
lt1/dk1-swapped relative to document order and that dk1/lt1 are usually
`<a:sysClr>` (windowText/window) with a `lastClr` fallback**. (b)
`Workbook.resolveColor(color: Color): string | undefined` — applies the MS tint algorithm
(HSL-luminance transform, not linear RGB) and returns 8-digit ARGB; `undefined` for
`{auto}` or when no theme part exists. Raw model unchanged. (c) The bridge carries the
SOURCE theme part verbatim: `WorkbookInput.themeXml?: string` (opaque, trusted, must be
`isXmlSafe`-clean and non-empty); when present the writer emits it instead of
`DEFAULT_THEME_XML` — removing the custom-theme flattening from the drop list.
**Design decisions (made).** Resolution is a METHOD on Workbook (needs the theme), not on
the style objects (they stay plain data). Tint rounds exactly as Excel does (round to
integer RGB after the Lum transform) — validate against openpyxl's `theme` handling on a
real custom-theme fixture, cell by cell.
**Tasks**
- [x] theme parser + resolveColor + tests (sysClr, srgbClr, tint± cases, index swap).
- [x] bridge/writer theme carry + byte-identical theme part round-trip test.
- [x] custom-theme fixture (`openpyxl-customtheme.xlsx` — accent1 recolored FF0000) + provenance.
- [x] adversarial review. (README resolved-color example deferred to the 0.5 release-prep docs pass.)
**Acceptance.** `resolveColor` matches the reference on every themed cell of the fixtures; a
custom-theme file round-trips with its theme byte-identical.
**Landed (uncommitted, awaiting owner approval).** `ooxml/theme.ts`: `parseTheme` (12-slot table
with the dk/lt index swap 0→lt1/1→dk1/2→lt2/3→dk2, sysClr lastClr fallback) + `resolveTint`
(**Excel's Win32 integer HLS, HLSMAX=240** — the float-HSL variant other libs use is off by ~1 per
channel; validated against 96 independent reference vectors AND real Excel swatches, grayscale +
accents exact) + `resolveColor`. `Workbook.resolveColor(color)` → 8-digit ARGB (rgb passthrough,
`{theme,tint}` resolved with an out-of-range/non-finite-tint clamp; `undefined` for auto/indexed/
no-theme/out-of-range); the reader loads `theme1.xml` and exposes `themeXml`. `WorkbookInput.themeXml`
carried by the bridge + emitted by the writer (validated non-empty + isXmlSafe, single-read) instead
of `DEFAULT_THEME_XML` — **custom themes now round-trip byte-identically; the last color flattening
is removed.** openpyxl reads our custom-theme rewrite warnings-as-errors. Gate: biome 0 / tsc 0 /
**418 tests** / byte-identity 7/7 vs pre-F5.3 (incl. theme-color case) / all emitted parts
well-formed. **Note (F5.3 scope decisions):** `openpyxl` has NO theme-resolution API, so the
independent oracle is a from-spec reference impl + Excel swatches, not openpyxl; the custom-theme
fixture is a post-processed openpyxl body (no headless Excel/LibreOffice locally) — a real-producer
custom theme is a welcome future upgrade; `{indexed}` palette resolution deferred (raw index kept).
**Adversarial review: 1 CONFIRMED bug fixed + regression-pinned** — a present-but-EMPTY (0-byte)
theme part read back as `themeXml:""`, which the bridge carried into the writer's non-empty check →
typed throw on a file the reader accepted (a round-trip regression). Fixed at the reader: an empty
theme part normalizes to `undefined` (empty ≡ absent), so it degrades to the built-in theme on
rewrite; the writer's strict `""` rejection stays correct for genuine direct-caller mistakes.

### F5.4 — Formula text: read → bridge → write (no evaluation) ☑
**Context.** Formulas are the largest remaining fidelity gap: the reader keeps only the
cached value; SheetJS/ExcelJS/openpyxl all preserve formula text. Evaluation stays M8 —
this feature is text fidelity only.
**Scope (in).** Reader: `Worksheet.formula(ref): string | undefined` — a lazy dedicated
scan (the `parseCellStyles` idiom) over `<f>` elements. Plain formulas verbatim. **Shared
formulas** (`t="shared"`, master carries text + `ref`, dependents carry only `si`):
dependents return the TRANSLATED text — relative A1 tokens shifted by the row/col delta
from the master, `$`-absolute parts pinned (a pure-string translator over the existing
a1.ts primitives; openpyxl's `Translator` is the reference behavior). **Array formulas**
(`t="array"`): text verbatim; the reader also exposes nothing extra (the ref lives in the
write model below). Writer: the object cell form gains `formula?: string` (+ optional
cached `value`); emits `<c r s?><f>…</f><v>cached</v></c>`; array masters emit
`<f t="array" ref="…">`. An error-typed cached value writes `t="e"` — so error cells WITH
a formula stop flattening to text. Validation: `isXmlSafe`, non-empty, ≤ 8192 chars
(Excel's formula ceiling — a shared bound in ooxml), no leading `=` (stored form).
**Scope (out).** Evaluation (M8); data-table formulas (`t="dataTable"`) degrade to
cached-value-only, documented; defined-name resolution.
**Design decisions (made).** `Cell` stays a value union — formula is an orthogonal, lazy
accessor, exactly like `style(ref)`. The bridge attaches `formula` whenever
`formula(ref)` resolves; cached value always written so non-recalculating consumers still
see data. The shared-formula translator is its own module with exhaustive unit tests
(absolute/relative/mixed refs, ranges, cross-sheet refs untouched, quoted sheet names,
strings containing ref-lookalikes NOT translated — tokenize, don't regex-replace blindly).
**Tasks**
- [x] `<f>` scan (`parseFormulas`) + shared-formula translator (`ooxml/formula.ts`) + array handling;
      reader tests + shared-formula fixture (`shared-formula.xlsx`).
- [x] writer emission + validation (no date1904 interplay — formula is text).
- [x] bridge carry; corpus snapshot gains `formula`; drop-list shrinks to bare error cells only.
- [x] openpyxl agreement: `data_only` both ways reads our formulas + cached values; adversarial review.
**Acceptance.** basic.xlsx's cached formula (`E1 = B1*2`) round-trips as a live formula;
the corpus property holds with formulas in the snapshot.
**Scope refinement (owner-notified).** ONE SHARED MODEL kept intact: `Worksheet.formula(ref)` returns
`string`, so the writer accepts `formula?: string` (symmetric). Consequences vs the original sketch:
array formulas carry the master's text as a PLAIN formula (array-ness/ref not exposed — a write-only
`t="array"` form would be a parallel "writer flavor" the invariants forbid); an error-typed cached
value on a formula cell writes as its STRING text (not `t="e"` — Excel recomputes the real error on
open, so the formula, not the stale cached error, is what matters); `dataTable` degrades to
cached-value-only. All documented in from-workbook.ts.
**Landed (uncommitted, awaiting owner approval).** `ooxml/formula.ts`: `translateFormula` (a linear
TOKENIZER — strings/quoted-sheets/whole-col & whole-row ranges/cell refs/identifiers — shifting
relative refs incl. the cell part of cross-sheet refs, pinning `$`, `#REF!` off-grid; matches
openpyxl's `Translator` on **32 reference vectors**) + `MAX_FORMULA_LEN` (8192). `Worksheet.formula(ref)`
+ `parseFormulas` (plain/shared-translated/array/dataTable). Writer `{ formula, value?, style? }`
cell + validation (non-empty, XML-safe, no leading `=`, ≤8192, single-read TOCTOU; string result
`t="str"`). Bridge carries formula + cached value. Gate: biome 0 / tsc 0 / **446 tests** /
byte-identity 8/8 vs pre-F5.4 / openpyxl `data_only` both ways / translator LINEAR on adversarial
input (500k-quote formula in 1ms). **Adversarial review: 3 CONFIRMED bugs fixed + regression-pinned**
— (1) whole-column/row refs (`A:A`, `1:1`) were left UNSHIFTED (silent formula corruption on a common
construct) → added range tokens to the translator, openpyxl-matched; (2) the shared second pass was
O(dependents × master-length) — a file-size-quadratic DoS on a hostile file → cap master text at
`MAX_FORMULA_LEN` (over-long masters degrade), re-verified 30k deps of an 18k-char master read in
40ms; (3) `parseFormulas` lacked the row assembler's `inSheetData` guard, so a stray `<c><f>` in an
`oleObjects`/`AlternateContent` block fabricated a formula on a real value cell → added the guard.
(A sibling latent bug — `parseCellStyles` lacks the same guard — is flagged as a separate follow-up,
out of F5.4 scope.)

### F5.1 — Constant-memory streaming writer ☑
**Context.** `writeXlsx` buffers every part, then zips (`writeZip(parts)`); fine to ~1M
cells, wrong for export jobs. The reader has streamed since F2.2 — this is the writer
mirror. Lands after F5.2–F5.4 so it reuses the final emission helpers.
**Scope (in).** `streamXlsx(workbook, options?): ReadableStream<Uint8Array>` where each
sheet's `rows` may be `Iterable | AsyncIterable` of row arrays (`CellInput[]`). Zip layer
gains a streaming path: local headers with **bit-3 data descriptors** (CRC + sizes
computed incrementally, written after each entry), central directory buffered (tiny) and
emitted at the end. Part order: worksheets stream first (styles intern as rows flow),
then styles/theme, workbook.xml, rels, `[Content_Types].xml` (order inside a zip is
irrelevant to OPC — everything resolves through rels).
**Scope (out).** Streaming READ of inputs other than rows (geometry/metadata stay
upfront values); re-entrant/parallel sheet streams (sheets serialize in order).
**Design decisions (made).**
- Streamed sheets omit `<dimension>` (optional per schema; bounds unknowable upfront) —
  documented, and the buffered writer's bytes are untouched (its own golden pins hold).
- Geometry (cols/sheetViews) precedes sheetData → written from the upfront sheet fields;
  mergeCells/hyperlinks/comments follow sheetData → validated upfront, emitted after the
  last row. Same validation code paths as buffered — no second validator.
- Backpressure: rows are pulled only when the consumer pulls (ReadableStream `pull`), so
  an async row source (DB cursor) is never outpaced.
- Determinism holds (no timestamps; fixed DOS date as in `writeZip`). Byte-identity with
  the BUFFERED writer is NOT a goal (descriptor layout differs by design) — equivalence
  is asserted through the reader: same input → identical reader snapshot.
**Tasks**
- [x] zip streaming entries (`stream-zip.ts`, bit-3 data descriptors) + incremental CRC
      (`crc32Update/Init/Final`); unit tests vs `openZip` and `unzip -t`.
- [x] worksheet row-stream serializer (`sheet.ts` streamWorksheet, reuses renderCell/emitters);
      `streamXlsx` + `StreamWorkbookInput`/`StreamSheetInput`/`StreamRows` + facade export.
- [x] equivalence property (buffered vs streamed → identical reader snapshot across corpus-shaped
      inputs); memory harness `scripts/stream-memory.mjs`.
- [x] example `examples/08-streaming-write.mjs` (DB-cursor async source); adversarial review.
**Acceptance.** 500k-row export completes with flat memory; openpyxl + Excel open the
streamed file; openXlsx snapshot equals the buffered writer's for identical input.
**Landed (uncommitted, awaiting owner approval).** `writer/stream-zip.ts` (streaming ZIP with bit-3
data descriptors + incremental CRC + streaming CompressionStream, buffered central dir), `writer/
stream.ts` (`streamXlsx` → ReadableStream, pull-based backpressure), `sheet.ts` `streamWorksheet`
(omits `<dimension>`), `writer/parts.ts` (OPC builders EXTRACTED from writeXlsx and SHARED — buffered
stays byte-identical). `crc32Update/Init/Final`. Gate: biome 0 / tsc 0 / **454 tests** / buffered
byte-identity 9/9 vs pre-F5.1 / streamed 50k-row async-cursor file passes `unzip -t` + openpyxl
warnings-as-errors / equivalence (streamed == buffered reader snapshot) / **memory harness: FLAT
~5.5MB live heap for 100k..1M rows** (constant memory, independent of file size). New example +
harness script. **Adversarial review: 4 CONFIRMED bugs fixed + regression-pinned** (1 refuted —
a benign validation error-message reorder from the parts.ts refactor): (1/2/3, HIGH, three lenses)
an early consumer `cancel()` LEAKED the row source — `deflateChunks`' detached pump wedged on
compressor backpressure forever, so a DB cursor's `finally` never ran and the CompressionStream
leaked → `deflateChunks` now drives the source via an explicit iterator and a `finally` that cancels
the reader, returns the source, and settles the pump (re-verified: cursor closed, 0 post-cancel
pulls); (4, MEDIUM) `streamXlsx`/`writeXlsx` multi-read `sheet.name`, so a flip-getter could slip a
forbidden name past validation into workbook.xml → `validateSheetMeta` now returns the validated
names and `workbookXml` emits from them (single-read), fixing both writers, byte-identity preserved.

### F5.5 — Benchmark harness + published numbers ☑
**Context.** "Fast" is currently an architectural claim (0.4 competitive analysis). The
1.0 gate needs published benchmarks; the harness comes now so every later feature can be
measured against a baseline.
**Scope (in).** A private `packages/bench` workspace (dev-only; competitor libs as
devDependencies THERE — the zero-dep principle applies to published packages only):
generated workloads (10k / 100k / 1M cells; strings-heavy, numbers-heavy, styled), read
and write, wall-time + peak heap, N iterations with warmup, machine-info stamp. Compare:
openjsxl (buffered + streamed), `exceljs@4.4.0`, `xlsx@0.18.5` (the npm-installable
SheetJS), and out-of-band reference numbers for python `openpyxl`/`python-calamine` from
a companion script. Output: a reproducible `docs/benchmarks.md` table (`pnpm bench`).
**Scope (out).** CI-run benchmarks (noise); micro-benchmarks of internals.
**Tasks**
- [x] harness + workload generator (deterministic datasets; materialized for buffered writers,
      lazy generator for the streamed writer); per-cell isolated worker with warmup/median +
      peak-RSS sampling; markdown reporter.
- [x] first published `docs/benchmarks.md` + README link with date + hardware note.
**Acceptance.** `pnpm bench` reproduces the table end-to-end on a clean checkout; README
claims match the published numbers (no unmeasured "fast" claims anywhere in docs).
**Landed (uncommitted, awaiting owner approval).** New private `packages/bench` (`@openjsxl/bench`,
never published — ExcelJS/SheetJS are devDeps THERE so the shipped packages stay zero-dep). `pnpm
bench` builds the real `openjsxl` bundle, authors read fixtures once with ExcelJS (a neutral
shared-strings producer, cached under gitignored `.cache/`), then runs every `(library, op,
workload, size)` cell in its **own `node --expose-gc` child process, strictly serially** — the one
design choice that keeps the numbers honest (no cross-library heap/CPU contention). Warmup +
median wall-time + peak RSS; every writer gets the identical dataset, every reader parses the
identical file and materializes each cell into a checksum sink; SheetJS writes with
`compression:true`; SheetJS styled-write is reported `—` (Pro-only) not a fake fast number;
openjsxl's streamed writer is fed a lazy row source (the real streaming case). Companion
`py/bench_py.py` (openpyxl + python-calamine, subprocess-isolated, same fixtures) merges via
`--render-only`. **Published `docs/benchmarks.md`** (M2 Pro, Node 24): at 1M cells openjsxl reads
~2–3× faster than ExcelJS/SheetJS at ~⅓ the memory, writes ~4× faster than ExcelJS, and the
streamed writer holds roughly flat ~130–180 MB vs 0.5–0.7 GB buffered / 1.4–1.8 GB ExcelJS.
Gate: biome 0 / tsc 0 / **454 tests** (no production code touched — a new private package + docs +
a README link, so byte-identity/corpus/openpyxl-writer gates are N/A). Streamed writer holds roughly
flat ~95–125 MB at 1M cells vs ~0.4–0.55 GB buffered / ~1.5–1.8 GB ExcelJS. **Adversarial fairness
review:** 4-lens workflow, 3 finders rate-limited (spend cap → UNVERIFIED, re-audited by hand:
equal-work PROVEN via identical cross-reader checksums, zero-dep intact, neutral fixture producer);
1 CONFIRMED (measurement) fixed — peak RSS now measured on the cold run only (warmed-iteration
residue over-reported it ~8–13%), dead `baselineRss` removed, doc wording corrected. **M5 COMPLETE
(F5.2 ✓ F5.3 ✓ F5.4 ✓ F5.1 ✓ F5.5 ✓).** Next: 0.5 release prep (README fidelity table + examples
updated BEFORE any version bump, per CLAUDE.md #4).

---

## M6 — Images (v0.6)

**Theme.** Pictures are the last fidelity gap a user can *see*: every real invoice or report
has a logo, and today the reader drops drawings entirely and the bridge silently loses them.
M6 closes it — drawingML picture read, anchored picture write (both writers), bridge carry.

**Re-scope vs the original outline (owner-approved with this section): the native lane is
DEFERRED, not built.** The original M6 bundled the optional `@openjsxl/native` napi-rs/WASM
`calamine` backend, "justified (or not) by the F5.5 numbers". The numbers came back *not*:
pure-TS openjsxl reads 1M cells in ~0.70 s — within ~1.5× of native `python-calamine`
(~0.46 s) and 2–3× faster than every JS incumbent (docs/benchmarks.md, 2026-07-03). A
prebuilt-binary release matrix is permanent maintenance to buy back fractions of a second,
and it serves exactly the runtimes (browser/edge) that are the differentiator worst. The
zip/xml interface stays swappable; revisit only if a future workload shifts the math.

**Standing gates for every M6 feature** (see CLAUDE.md for the full contract): biome by
exit code + tsc + vitest green; byte-identity for input not using the new feature (in-tree
golden pins + the worktree recipe); the corpus property test extended to newly-carried
state; openpyxl cross-validation both directions; adversarial review (finders + refuting
verifiers; UNVERIFIED ≠ refuted — re-verify dead agents' claims empirically) with confirmed
findings fixed and pinned before the commit is proposed. Probes stay in the scratchpad.

**Dependency order for hand-off:** F6.1 (the per-sheet wiring dedup) lands FIRST — the M5
analysis flagged that the rel/part wiring already exists in 2–3 hand-synchronized copies,
and images add a fourth per-sheet part family; pay the debt before growing it. Then
F6.2 (read) → F6.3 (write) → F6.4 (bridge + corpus). One feature per session, each
proceed-gated: "proceed" → implement → gates → adversarial review → report + commit
message → WAIT for the owner's "commit".

### F6.1 — Per-sheet part wiring dedup (pre-image refactor) ☑
**Context.** The hyperlinks→comments→vmlDrawing rel block and the comment/VML part-path
conventions are spelled independently in `worksheetXml` (writer/sheet.ts ~779–799),
`streamWorksheet` (~926–944), the buffered part loop (writer/workbook.ts ~98–104), and the
streaming part loop (writer/stream.ts ~67–83); `contentTypesXml` takes a `needVml` flag both
callers derive identically. All copies agree today by hand; F6.3 would make it four.
**Scope (in).** One shared per-sheet builder (in writer/parts.ts or sheet.ts) that owns the
rel ordering, rId allocation, body element refs (`legacyDrawing`, later `drawing`), and the
part-name conventions (`xl/comments{N}.xml`, `xl/drawings/vmlDrawing{N}.vml`, later
`xl/drawings/drawing{N}.xml` + `xl/media/*`); both writers consume it. Derive `needVml`
inside `contentTypesXml` from `commentSheets`.
**Scope (out).** The zip-primitive duplication (`u16`/`u32`/DOS constants in zip.ts vs
stream-zip.ts) — separate, optional cleanup; do NOT bundle it here.
**Design decisions (made).** Pure refactor: **zero byte change** — the golden pins and the
full byte-identity recipe are the acceptance gate, not a nice-to-have. No new types on the
public surface.
**Tasks**
- [x] Extract the shared per-sheet parts/rels builder; both writers consume it; delete the
      copies. `needVml` derived, parameter dropped.
- [x] Gates + byte-identity (full pin set + worktree recipe if in doubt) + adversarial
      review focused on emission-order/rId regressions.
**Acceptance.** Identical bytes for every input in the pin set; suite green; the wiring
exists in exactly one place.
**Landed (committed `4485219`).** Found the dedup HALF-done: `sheetRelPlumbing` existed and
`worksheetXml` used it, but `streamWorksheet` still inlined the rel block and the side-part NAMES
were spelled in both part-loops. `streamWorksheet` now calls `sheetRelPlumbing`; new
`sheetSideParts(sheetIndex, side)` in parts.ts owns the side-part OPC names
(`_rels/sheetN.xml.rels`, `commentsN.xml`, `drawings/vmlDrawingN.vml`) in one place (rels→comments→
vml order), consumed by both writers; `contentTypesXml` derives `needVml` from `commentSheets`
(param dropped); streaming derives `commentSheets` upfront from `prepared`. **F6.3 image parts
(drawingN.xml + drawing rels) extend `SheetSideParts` + `sheetSideParts` here.** Gate: biome 0 /
tsc 0 / **458 tests** (+2 F6.1 guards: both writers emit the identical OPC part set; sheet1 rels
stay ordered hyperlink/comments/vmlDrawing) / **byte-identity 10/10** vs the pre-F6.1 worktree
(`369f36f`) incl. every side-part path / streamed side-parts read back + openpyxl warnings-as-errors
clean. Adversarial review (1 focused agent): NO issues (5 checks). Pure refactor, zero behavior change.

### F6.2 — Picture read: `Worksheet.images` ☑
**Context.** In `.xlsx`, pictures hang off the worksheet's `<drawing r:id>` →
`xl/drawings/drawingN.xml` (spreadsheetDrawing `xdr:` namespace) → `xdr:oneCellAnchor` /
`xdr:twoCellAnchor` each containing `xdr:pic` → `xdr:blipFill` → `a:blip r:embed` → the
drawing part's OWN rels → `xl/media/imageK.<ext>`. Resolve through the relationship graph,
never by filename. Mime comes from `[Content_Types].xml` (`Default` by extension, rarely an
`Override`).
**Scope (in).** Lazy `Worksheet.images: readonly SheetImage[]` (the `mergedCells` idiom —
parsed on first access, cached). Shared model (this IS the writer input in F6.3):
`SheetImage = { anchor: ImageAnchor; bytes: Uint8Array; mime: string; name?: string }`;
`ImageAnchor = { from: AnchorPoint; to?: AnchorPoint; ext?: { cx: number; cy: number };
editAs?: 'twoCell' | 'oneCell' | 'absolute' }`; `AnchorPoint = { col: number; row: number;
colOff?: number; rowOff?: number }`. `to` present ⇔ twoCellAnchor; `ext` present ⇔
oneCellAnchor.
**Scope (out).** `absoluteAnchor` pictures (skipped, documented); non-picture drawing
objects — shapes, charts, group/graphic frames (skipped, documented); chartsheets; picture
effects/crop (`srcRect`, filters) — dropped, documented.
**Design decisions (made).**
- **Raw anchor model** — the colors precedent: only the raw shape round-trips faithfully.
  EMU offsets/extents kept verbatim as integers (no px helpers; 914 400 EMU/inch, 9 525
  EMU/px @96dpi goes in the jsdoc, not the API).
- **`col`/`row` are 1-based** in the public model (consistent with the whole API); OOXML
  anchors store 0-based — convert at parse and at emit, test the fencepost explicitly.
- **Bytes are opaque.** Never decode image data. Media parts inflate on demand, cached per
  part (two pics sharing one media part share one cached `Uint8Array`).
- **Tolerant reader degrades:** a missing media rel/part, an unresolvable r:embed, or an
  absoluteAnchor pic ⇒ that picture is skipped, never a throw; a structurally corrupt zip
  part still fails typed as everywhere.
- Mime degraded from the extension when content-types has no entry; unknown extension ⇒
  `application/octet-stream` (reader tolerant; the WRITER's allowlist is the strict side).
**Tasks**
- [x] `ooxml/drawing.ts` — parse drawingN.xml anchors + pics via the tokenizer (never a
      DOM); resolve blip r:embed through the drawing's rels.
- [x] `Worksheet.images` lazy accessor + media caching; export `SheetImage`/`ImageAnchor`/
      `AnchorPoint` types from the core index.
- [x] Fixtures: `openpyxl-images.xlsx` (openpyxl + Pillow in the scratchpad venv: one
      oneCellAnchor PNG, one twoCellAnchor JPEG, two pics sharing one media part) with
      provenance per data/README.md; a hand-built fixture for the degrade cases
      (absoluteAnchor, missing media part) via the fixtures builder.
- [x] Tests: exact byte round-out (digest vs the source image), anchors (1-based fencepost
      pinned), mime, shared-media caching, every degrade case; openpyxl agreement on
      anchor cells + image bytes.
**Acceptance.** Fixture images read back with byte-identical content and correct anchors;
degrades never throw; suite green.
**Landed (uncommitted, awaiting owner approval).** `ooxml/drawing.ts` `parseDrawing` (pure SAX
parse: one/twoCellAnchor pictures → `{embed, name?, from, to?, ext?, editAs?}`; 0-based→1-based;
skips absoluteAnchor, non-pic anchors, and grpSp-grouped pictures) + `mimeForMediaPath`. Reader:
`Worksheet.images(): Promise<readonly SheetImage[]>` — **an async method, not the scoped sync
getter** (deviation, owner-notified): image bytes need async decompression, and eagerly reading all
media at open would break the reader's "a sheet you never touch costs a decompression, not a parse"
principle. It is lazy (resolves on first call via a per-sheet thunk built in openXlsx, cached) and
matches the existing async `rows()`; the shared `SheetImage` TYPE is still identical to the writer
input. `loadSheetImages` resolves `/drawing` rel → drawing part → parseDrawing → drawing rels →
media (read once per part, shared buffer); every miss degrades to skip, never throws. Types
`SheetImage`/`ImageAnchor`/`AnchorPoint` exported. Fixtures: real `openpyxl-images.xlsx`
(openpyxl 3.1.5 + Pillow: twoCellAnchor JPEG + oneCellAnchor PNG) validates the parser on genuine
drawingML (unprefixed anchors, package-absolute media targets); crafted `images-edge.xlsx`
(packParts, now binary-capable) covers shared-media + every degrade path. Gate: biome 0 / tsc 0 /
**473 tests** (+15) / openpyxl independent-reader cross-check agrees on bytes (SHA) + anchors /
byte-identity N/A (reader-only). **Adversarial review (2 lenses): 1 CONFIRMED defect fixed + pinned**
— a picture's `spPr/a:xfrm/a:ext` (emitted by real Excel/LibreOffice) overwrote the anchor's own
`<ext>` (mis-sizing one-cell pics; injecting a model-violating `ext` onto two-cell anchors) → the
`ext` branch now captures only the anchor-level ext (before `<pic>`, with cx); the grpSp observation
was also fixed (grouped pics skipped, `cNvPr` name gated to the pic). Hostile-input lens: no throw/
hang/superlinear (1M anchors ~1.4s linear; no path traversal; all degrades skip).

### F6.3 — Picture write (both writers) ☑
**Context.** The writer mirror: `SheetInput.images?: readonly SheetImage[]` — the reader's
exact shape (one shared model, no writer flavor).
**Scope (in).** Per sheet with images: `xl/drawings/drawingN.xml` (+ its rels part) and the
worksheet `<drawing r:id>` element — schema position: **before `legacyDrawing`** in the
CT_Worksheet sequence (…mergeCells → hyperlinks → … → drawing → legacyDrawing); media parts
`xl/media/imageK.<ext>` with **workbook-level dedup**; content-types `Default` entries per
used extension. Both writers (buffered + streaming) — images are upfront metadata like
comments, so the streaming writer emits them identically; media parts are plain byte parts.
**Scope (out).** Generating anchors from pixel sizes (callers give EMU); image re-encoding
or dimension sniffing (bytes are opaque — a caller wanting "natural size" measures the
image themselves); absoluteAnchor write.
**Design decisions (made).**
- **Validation (typed `invalid-input`, naming the sheet + image index):** mime allowlist
  `image/png` | `image/jpeg` | `image/gif` (extension derived from mime, never from data);
  `bytes` a non-empty `Uint8Array`; anchor cols/rows within the shared grid bounds
  (MAX_ROW/MAX_COL from ooxml/a1.ts); EMU offsets/extents non-negative safe integers
  (cap at 2^31−1 — the XML schema's int); exactly one of `to` (twoCellAnchor) or `ext`
  (oneCellAnchor) present; `isPlainRecord` + unknown-key rejection; **single-read TOCTOU**
  on every property, incl. `bytes` (read the reference once; a getter must not swap the
  buffer between validation and packing).
- **Media dedup is deterministic:** identical bytes ⇒ one `xl/media/` part, numbered by
  first-occurrence order; equality by length-grouped byte-compare (no hashing, no
  randomness, linear in total media bytes).
- **Unused emits nothing:** an imageless workbook produces exactly the pre-F6.3 bytes
  (byte-identity gate), no drawing/media parts, no content-type entries.
- Drawing XML from minimal templates with escaped/validated substitutions, like VML (F5.2);
  `xdr:pic` gets the minimal required children (nvPicPr with a deterministic id/name,
  blipFill, spPr with a bare prstGeom) — copy openpyxl's emission as the reference shape.
**Tasks**
- [x] Shared model types finalized; validation module; media registry (dedup + numbering).
- [x] drawingN.xml emission + drawing rels + worksheet element + content types, wired
      through the F6.1 shared per-sheet builder for BOTH writers.
- [x] Tests: write→read deep-equal (bytes byte-identical, anchors exact); byte-identity for
      imageless input; streamed == buffered reader snapshot with images; validation
      rejections (each rule); dedup (same bytes twice ⇒ one media part; almost-equal ⇒
      two); TOCTOU flip-getter pins.
- [x] `unzip -t` + openpyxl reads OUR images (bytes + anchors) warnings-as-errors.
**Acceptance.** Excel/LibreOffice (openpyxl proxy) shows the written picture at the right
cell; round-trip lossless; imageless output byte-identical.
**Landed (uncommitted, awaiting owner approval).** `SheetInput.images?: readonly SheetImage[]` (the
F6.2 reader shape — one shared model). New `writer/images.ts` `createMediaRegistry` (workbook-level
media dedup, deterministic FNV-1a content hash keyed by `(bytes, ext)` with byte-compare on
collision — deviates from the scope's "byte-compare, no hashing" to avoid **O(n²) on bridge-fed
attacker-controlled image counts**, a CLAUDE.md invariant). `imageParts` (sheet.ts) validates each
picture (mime allowlist png/jpeg/gif via **Object.hasOwn**; non-empty Uint8Array bytes single-read;
exactly one of `to`/`ext`; grid bounds; EMU 0..2³¹−1; isPlainRecord+unknown-key at every level; TOCTOU
single-read incl. bytes + to/ext) and builds `drawingN.xml` + its rels. `sheetRelPlumbing` emits the
`<drawing r:id>` element **before** `<legacyDrawing>` (CT_Worksheet order); `sheetSideParts` emits the
drawing + drawing-rels parts; `contentTypesXml` adds drawing Overrides + media Defaults (empty when
no images). Both writers create ONE media registry and emit `xl/media/imageK.<ext>` workbook-level.
`packParts` gained binary support (unrelated fixtures helper). Gate: biome 0 / tsc 0 / **492 tests**
(+19 image-write) / **byte-identity 10/10** imageless vs pre-F6.3 worktree / write→read round-trip
(bytes byte-identical, anchors exact) / dedup (3 pics→2 parts; cross-mime→2 parts) / streamed ==
buffered / `unzip -t` + openpyxl reads our output warnings-as-errors. **Adversarial review (2 lenses):
2 CONFIRMED defects fixed + pinned** — (1) mime allowlist bypass via `Object.prototype` keys
(`mime:"constructor"` → garbage media part + `ContentType="undefined"`) → gated on `Object.hasOwn`;
(2) dedup key ignored the extension, so identical bytes as png+jpeg wrote one part and the second
picture's rel pointed at a nonexistent file (silent image loss) → dedup identity now includes `ext`.
Everything else checked sound (name/editAs escaping, EMU validation, TOCTOU, cross-sheet dedup, FNV
false-merge, schema order, hostile scale). **Note (deviation, owner-notified):** deterministic hash
dedup instead of the scoped pure byte-compare, for adversarial-input safety.

### F6.4 — Bridge carry + corpus + example ☑
**Scope (in).** `workbookToInput` carries `sheet.images` per sheet (absent when none —
byte-identity path preserved); the corpus property snapshot gains images (anchor + mime +
a bytes digest, not raw bytes); example `10-images.mjs` (write a logo'd report, read it
back, round-trip it); fidelity docs updated in the README **at 0.6 release prep** (not
here) — the drop-list line becomes "error cells; absolute-anchored/non-picture drawings".
**Tasks**
- [x] Bridge attachment + corpus snapshot extension (every fixture with pictures
      round-trips losslessly or fails typed).
- [x] Example 10 + examples README/package.json wiring; all examples green.
- [x] Full-milestone adversarial review round (cross-feature, the M5-analysis style) +
      fixes pinned. Release prep (README fidelity table + examples) then 0.6 bump ONLY at
      the owner's explicit request. *(release prep + bump deferred to owner request)*
**Acceptance.** read → bridge → write → read gives identical images across the corpus;
the M6 drop-list documents exactly: error cells, absolute-anchored pics, non-picture
drawing objects, picture effects.

**Landed.** `workbookToInput` (from-workbook.ts) now `await`s `worksheet.images()` per
sheet and attaches `images` only when non-empty — an imageless workbook keeps the exact
pre-F6.3 WorkbookInput and the exact same bytes (the emitters are untouched). The corpus
property snapshot (bridge-styles.test.ts) gains `images` compared by anchor + mime + name
+ a content **digest** (length + FNV-1a, never raw bytes), so `openpyxl-images.xlsx` (jpeg
+ png) and `images-edge.xlsx` (shared-media + degrade paths) now round-trip through the
snapshot (2 images each). New example `10-images.mjs` (self-contained real PNG; write →
`images()` → bridge round-trip) + examples `package.json`/README wiring; all 10 examples
green vs built dist. **Full-milestone adversarial review (3 lenses: round-trip/bridge,
hostile-input, spec/algorithm): 1 CONFIRMED defect fixed + pinned** — a malformed drawing
part could make `parseDrawing` return an anchor with NEITHER `to`/`ext` (one-cell missing
its `<ext>`) or BOTH (stray `<ext>` on a two-cell), which the writer's "exactly one of"
rule then rejects, throwing typed and nuking the WHOLE read→write rewrite (and falsifying
the reader's own "returns valid writer input" contract). Fixed at the reader: the anchor
ELEMENT names its shape, so `parseDrawing` keeps only the field the kind calls for
(twoCellAnchor→`to`, oneCellAnchor→`ext`), recovering a stray-field picture and dropping
one whose kind-required field is absent — the same degradation as a missing blip/media
part. Everything else checked sound (cross-sheet drawing/media numbering by sheet index vs
workbook-level media dedup; images+comments+hyperlinks coexisting with correct rIds and
`drawing` before `legacyDrawing`; buffered == streamed; `editAs` orthogonality carried
losslessly; exotic-mime images fail TYPED not bare — same shared-bounds "reject" path as
oversized EMU/off-grid cells). Gate: biome 0 / tsc 0 / **495 tests** (+1 drawing shape
regression, +2 bridge picture round-trip) / imageless byte-identity (emitters untouched;
pinned by the corpus unstyled + new imageless tests) / openpyxl reads our bridge-rewrite
AND a fresh write warnings-as-errors clean (2 images each). **Open design note (owner's
call, deferred):** the reader recognizes 10 media types but the writer's allowlist is
png/jpeg/gif — a real file with a bmp/tiff/webp/emf/wmf image round-trips as a TYPED
refusal (spec-compliant) rather than carrying; widen the writer (single-source the
mime↔ext map) or drop-and-document exotic images at 0.6 prep.

### 0.6 release prep (owner-approved) ☑
**Exotic-mime decision RESOLVED (owner picked: widen).** `MEDIA_MIME_TO_EXT` now covers the FULL
read set — png, jpeg, gif + bmp, tiff, webp, emf, wmf — so reader and writer are symmetric and
any real file with pictures round-trips; only a genuinely unknown type (the reader's
`application/octet-stream` fallback) still refuses typed. The reader's derived map keeps `jpg`/
`tif` as alternate SPELLINGS only; the writer's mime error message now enumerates from the map
(can't go stale). +5 pinned `it.each` round-trips (part ext + content-type Default + re-read
mime per type); the old "webp rejects" case became `image/avif` (genuinely unknown). **Docs
(P4 + fidelity):** root README (status → New in 0.6, `images()` in the reader block, new
"Pictures (0.6)" write section, fidelity table gains pictures row + absolute-anchor/non-picture
+ effects drop-entries, media renumber/ext-spelling flattening note, streamXlsx + benchmarks
"constant in ROWS" image caveat); core README (images() usage, images in the writer paragraph +
allowlist, stream caveat, SheetImage/ImageAnchor/AnchorPoint in exports); facade README (same
three); `images()` JSDoc gains the per-sheet media-cache note. ROADMAP 0.6 row wording already
accurate (ticks owner-managed, untouched). Gate: biome 0 / tsc 0 / **506 tests** /
**byte-identity 12/12** vs pre-widening (`9395ee9`) — trio output provably unchanged / openpyxl
warnings-as-errors clean on REAL Pillow-generated bmp+tiff+webp through our writer (emf/wmf not
PIL-decodable — openpyxl drops any image PIL can't read, incl. valid ones; identical OPC path)
/ `unzip -t` OK / all 10 examples green vs built dist. **Version bump NOT included** — separate
explicit owner request per CLAUDE.md #4 (owner runs push / `pnpm -r publish` / tag).

### M6 analysis follow-up (post-milestone, owner-approved) ☑
Post-milestone analysis (2 agents: scale+memory, debt+gaps; every finding re-verified
empirically) → one fix set. **P1 shared-bounds parity:** `parseDrawing` now CLAMPS a
producer's out-of-range anchor numbers into the writer's legal ranges (cols/rows into the
grid, offsets/extents into EMU 0..2³¹−1) instead of returning them verbatim — previously a
well-formed file with `colOff="-5000"` or `cx="2147483648"` read fine but the writer's
typed rejection made the WHOLE file un-rewritable (violating the "reader clamps or drops"
clause; the F4.5 geometry precedent). `MAX_EMU` moved to `ooxml/drawing.ts` as a shared
constant (the a1.ts pattern); the writer imports it. **P1b (found during the fix,
reproduced then pinned):** `mimeForMediaPath` did a prototype-chain lookup — a hostile
media part named `image1.constructor` surfaced the `Object` constructor FUNCTION as the
"mime" (breaking `SheetImage.mime: string`); now gated on `Object.hasOwn`, the reader-side
twin of the writer's F6.3 fix. **P2 single-sourced media types:** the THREE hand-synced
mime↔ext maps (reader `MIME_BY_EXT` 10-type, writer `MEDIA_MIME_TO_EXT`, content-types
`MEDIA_EXT_TO_MIME`) collapse to ONE canonical `MEDIA_MIME_TO_EXT` in `ooxml/drawing.ts`;
the reader's map and the content-types map now DERIVE from it (extra read-only types
listed once, reader-side). No behavior change — the writer allowlist stays png/jpeg/gif
pending the deferred exotic-mime decision. **P3 coverage:** +6 pinned tests — anchor
clamps; prototype-key mime; GIF end-to-end (part name + content-type Default + re-read);
XML-special-char picture-name round-trip; multiple `/drawing` rels on one sheet (crafted,
rel order); `TargetMode="External"` blip skip. **P4 doc items** (streaming "constant
memory in rows" image caveat, README fidelity tables, images() media-cache JSDoc note)
deferred to the 0.6 release-prep pass. Gate: biome 0 / tsc 0 / **501 tests** /
**byte-identity 12/12** (both writers × no-date, date, multi-sheet-sparse, empty-sheet,
comments+links, images-mixed vs the pre-change worktree at `f769a37`) — writer bytes
unchanged, so the F6.4 openpyxl validation stands transitively.

---

## M7 — More formats (v0.7)

**Theme.** Read the spreadsheets people actually have. Three read-only formats: `.ods`
(the LibreOffice / OpenDocument world), `.xlsb` (Excel's binary workbook — the default
for huge grids), and `.csv`/`.tsv` (the universal delimited-text export — what Google
Sheets, Excel "Save as CSV", and every database/BI tool produce, and the single most
common "user uploaded a spreadsheet" case). The writer stays `.xlsx`-only, so every new
reader is automatically a converter through the existing bridge
(`writeXlsx(workbookToInput(await openOds(bytes)))`) — which is the real-world job
("user uploaded a spreadsheet", in whatever dialect). Competitive frame: ExcelJS reads
none of these; SheetJS reads them all — this milestone closes the ingestion-breadth gap
while keeping the zero-dep / typed / tolerant story. **Legacy `.xls` (BIFF8) is
deliberately OUT** (owner call, 2026-07-10): it is the biggest lift in the milestone
(a whole CFB/OLE2 container + BIFF record layer + the SST `Continue`-split trap) for a
declining ~1997–2007 install base — deferred to M8+, revisited only if users ask.

**Milestone-wide decisions** (set here so features don't re-litigate):
1. **Read-only.** No ods/xlsb/xls writers. Conversion = bridge → `writeXlsx`, validated
   as everywhere: input Excel forbids (an over-long ods sheet name) fails TYPED on write,
   never silently renamed — reader degrades, WRITER rejects.
2. **One shared model.** Every format returns the SAME public `Workbook`/`Worksheet`
   surface; accessors a format doesn't support DEGRADE (`images()` → `[]`, `style()` →
   `undefined`, `themeXml` → `undefined`, …), never throw. Per-format fidelity is a
   documented matrix, not a type fork. The multi-format seam is internal (an F7.1 design
   task); existing xlsx classes, bytes, and tests stay untouched.
3. **Explicit entry points** — `openOds` / `openXlsb` / `openCsv` (+
   `detectSpreadsheetFormat`); no auto-dispatching mega-function in core: a browser
   bundle that only reads xlsx must not pay for the BIFF/CSV readers (tree-shakability
   verified in F7.4). `XlsxError` stays the one error type across formats.
4. **Fidelity scope = the data.** Values + types (string/number/date/bool/error), cached
   formula VALUES, sheet names/order, merges, `dimension`; visibility +
   `numberFormat(ref)` where the format hands them over anyway (xlsb/xls); hyperlinks
   where cheap (ods/xlsb). OUT for 0.7 (named, never silent): styles read, formula TEXT
   (xlsb/xls store token streams — a decompiler is its own feature; ods stores
   OpenFormula needing ref/separator translation), comments, images, geometry, and
   streaming variants of the new readers (`streamSheetRows` stays xlsx-only).
5. **UTF-8 in, conservative types.** `.csv`/`.tsv` is UTF-8 text (BOM stripped); other
   encodings are out (documented). Delimiter is auto-detected (comma / tab / semicolon —
   Excel uses `;` in some locales) with an override. Type inference is conservative:
   numbers and booleans only; **dates stay strings** (CSV date formats are locale-ambiguous
   — inferring them fabricates wrong values, the reader-degrades way). Encrypted inputs —
   ods manifest `encryption-data` — fail typed, never garbage-parse.
6. **Tolerant-reader rules span formats.** Shared grid bounds clamp/drop; adversarial
   input is a first-class case: ods `number-*-repeated` bombs, xlsb record-length lies +
   SST `Continue` splits, and CSV quote/newline pathologies (an unterminated quote making
   the rest of the file one field; a line with millions of columns) are each NAMED test
   cases, bounded in time and memory (the F4.4/F4.6 precedents).
7. **Oracle per format.** openpyxl reads none of these. Independent oracles: **python-calamine**
   (ods + xlsb) for values/types/dates, **pyxlsb** as a cheap xlsb second opinion, and
   **Python's stdlib `csv`** (RFC 4180 reference) for CSV quoting/newline behavior.
   Scratchpad venvs, as always. Crafted fixtures must ALSO parse in the oracles — a fixture
   only our reader accepts proves nothing.

**Standing gates** per feature: biome by exit code + tsc + vitest; adversarial review
(finders + refuting verifiers; UNVERIFIED ≠ refuted — re-verify empirically) with
confirmed findings fixed + pinned pre-commit; fixtures follow the data/README checklist
(real-producer + crafted-edge per format). The writer is untouched all milestone — the
in-tree golden pins must stay green; any feature touching SHARED code (a1 / dates / zip /
format-code logic) additionally runs the full byte-identity recipe.

**Dependency order:** F7.1 `.ods` (pure XML — reuses the tokenizer wholesale; defines the
multi-format seam with the least new machinery) → F7.2 `.xlsb` (the binary-record layer:
varint framing, RK, wide strings) → F7.3 `.csv`/`.tsv` (no container, no binary — a
robust delimited-text parser; **independent of F7.2**, could land anytime after the seam
exists, kept third for numbering) → F7.4 (detection, conversion corpus, docs, bench
lanes). One feature per session, each proceed-gated: "proceed" → implement → gates →
adversarial review → report + commit message → WAIT for the owner's "commit".

### F7.1 — `.ods` read: `openOds` ☑
**Context.** An ODF spreadsheet is a zip whose first entry is a STORED `mimetype`
(`application/vnd.oasis.opendocument.spreadsheet`) and whose sheets ALL live in one
`content.xml` (`office:spreadsheet` → `table:table` → `table:table-row` →
`table:table-cell`). Types are EXPLICIT — `office:value-type` plus a typed value
attribute — so there is no style-based date detection at all; dates/times are ISO-8601
strings, not serials. A different vocabulary over the same SAX tokenizer; the zip layer
is reused verbatim.
**Scope (in).** `openOds(source)` → the shared Workbook surface. Cell mapping:
`float`/`percentage`/`currency` → number (`office:value` authoritative — display text may
be locale-formatted); `boolean`; `string` (nested `text:p` paragraphs joined with `\n`;
`text:s [text:c]` / `text:tab` / `text:line-break` expanded; `text:a xlink:href` captured
as the cell's hyperlink); `date` (`office:date-value` → UTC-based `Date`, time-of-day
kept — the serialToDate convention); `time` (`office:time-value` ISO duration → number as
fraction of a day, documented). Formula cells read their CACHED value via the same typed
attrs. Repeats (`table:number-columns-repeated` on cells/covered cells,
`table:number-rows-repeated` on rows) materialize content-aware; merges from
`table:number-columns/rows-spanned` + `covered-table-cell`; sheet names/order from table
order. Container check tolerant: `mimetype` entry OR the content root's document class.
**Scope (out — named, documented).** Formula text (OpenFormula `of:=…`, `[.A1]` refs, `;`
separators — a translation feature of its own); data styles → `style(ref)` /
`numberFormat(ref)` read `undefined`; sheet visibility (lives in `settings.xml` config
items — sheets read visible); annotations (comments), images, geometry; flat `.fods`
(single-file XML, no zip); encrypted ods (typed `unsupported`).
**Design decisions (made).**
- **The repeat bomb is THE ods adversarial case.** Trailing empty repeated cells/rows
  (LibreOffice ends every sheet with a row repeated to the 2^20 grid edge) are DROPPED,
  not materialized; mid-content repeats materialize but clamp at the shared grid bounds
  (MAX_ROW/MAX_COL — reader degrades); total materialization is bounded by CONTENT,
  never by attacker-chosen repeat counts.
- One `content.xml` holds every sheet: the first sheet access parses the part ONCE and
  caches all sheets' tables (documented — per-sheet laziness can't exist inside a single
  part).
- The multi-format seam lands here (under decision 2's constraints). Leaning: extract
  public STRUCTURAL interfaces the existing xlsx classes already satisfy — no xlsx code
  motion, no consumer type break (`instanceof` was never documented API); final call in
  this feature's report.
**Tasks**
- [x] `ods/` dir: container check + `content.xml` SAX parse → per-sheet typed cell
      tables (repeats/covered/spans/links; caps + trailing-empty drop).
- [x] Multi-format seam + `openOds` + facade export; degrade wiring for unsupported
      accessors.
- [x] Fixtures: real-producer `.ods` (odfpy — no local LibreOffice; typed matrix + formula +
      merges + hyperlink + multi-sheet) + crafted edge/reject fixtures via the generator
      (repeat bomb, covered-cell merge, hidden/empty sheets; encrypted / wrong-mimetype /
      no-content rejects).
- [x] Tests: type matrix vs python-calamine cell-for-cell; repeat clamp/drop bounds;
      merges + links; every degrade case; xlsx suite untouched-green.
- [x] Adversarial review (hostile-input lens on repeats; model lens on the seam).
**Acceptance.** Fixture values/types/dates match python-calamine cell-for-cell; a crafted
repeat bomb parses (or rejects) in bounded time/memory; `openXlsx` behavior and the full
existing suite are untouched.

**Landed (uncommitted, awaiting owner approval → committed with this feature).** `openOds`
returns the SAME public `Workbook` as `openXlsx`. **Seam:** `Worksheet` is now a structural
INTERFACE (`types.ts`) — the `#private`-field class is nominally typed, so a parallel ODS class
can't be assignable to it; instead the xlsx class became `XlsxWorksheet implements Worksheet`
(pure rename, zero logic change), `OdsWorksheet implements Worksheet` is a plain data holder, and
the format-agnostic `Workbook` class is REUSED. `Row` moved to `types.ts`. Public change: `Worksheet`
was an exported class, now an exported interface — `instanceof Worksheet` breaks (undocumented, unused
in-tree). `ods/content.ts` is a pure SAX parser (typed values; repeats with an O(1) empty-tail drop;
covered-cell merges; `<text:a>` hyperlinks; a synthesized dimension); `reader/ods.ts` = `openOds` +
`OdsWorksheet` + container/encryption/mimetype checks (typed `XlsxError`). Fixtures: real
`odf-basic.ods` (odfpy, calamine-cross-checked) + crafted `ods-edge/-encrypted/-not-spreadsheet/
-no-content.ods`. Bridge (`workbookToInput` → `writeXlsx`) converts `.ods` → `.xlsx`. Gate: biome 0 /
tsc 0 / **528 tests** / cell-for-cell vs python-calamine / xlsx suite untouched-green.
**Adversarial review (4-lens workflow → refuting verifiers; 9 candidates → 2 CONFIRMED, 5 refuted,
both fixed + pinned):** (1) `parseOdsDate` shifted years 0–99 by +1900 (the `Date.UTC` two-digit-year
trap — `0050-01-15` read as 1950; calamine reads year 50) → `setUTCFullYear`; (2) `MAX_ODS_CELLS` was
PER-SHEET, so N repeat-bomb sheets materialized N×2M cells and OOM'd from a few KB of input → hoisted
to a DOCUMENT-WIDE `totalCells` budget. Bonus hardening: a grid-edge span that clamps to a single-cell
`"XFD1:XFD1"` range is dropped (not a merge). Probes stayed in the scratchpad.

### F7.2 — `.xlsb` read: `openXlsb` ☑
**Context.** Same OPC container as xlsx — `[Content_Types].xml` and `_rels/*.rels` are
XML, so the F1.4 relationship graph is reused VERBATIM — but workbook / worksheets /
sharedStrings / styles are BIFF12 binary parts (`.bin`): records framed as a 1–2-byte
varint id + 1–4-byte varint length (7 data bits per byte, high bit = continuation). Cell
semantics mirror xlsx exactly: shared-string indexes, style-index date detection (numFmt
CODES are the same strings — `isDateFormatCode` + the builtin-id table are reused), and
the 1904 flag in `BrtWbProp`.
**Scope (in).** `openXlsb(source)` → shared surface. New `biff/` dir for the BIFF12
primitives: bounded record stream (varint framing), wide strings (uint32 count +
UTF-16LE), RK decode (30-bit int-or-float + ÷100 flag).
Parsers: `workbook.bin` (BrtBundleSh → name/visibility/rel id; BrtWbProp → 1904),
`sharedStrings.bin` (BrtSSTItem — rich runs/phonetics skipped, text kept, the F1.5
convention), `styles.bin` (BrtFmt/BrtXF → numFmtId per xf, feeding date detection AND
`numberFormat(ref)`), `sheetN.bin` (cells BrtCellIsst/St/Rk/Real/Bool/Error/Blank +
BrtFmlaNum/Str/Bool/Error cached values; BrtWsDim → `dimension`; BrtMergeCell → merges;
BrtHLink → hyperlinks through the sheet's rels).
**Scope (out — named).** Formula text (BIFF12 stores token streams — a decompiler is its
own feature, deferred); `style(ref)` beyond numFmt; everything decision 4 already names.
**Design decisions (made).**
- The xlsx laziness idiom kept: open parses workbook + sst + styles; sheet parts parse on
  first access.
- The record reader is bounds-checked against its part: a lying record length fails
  typed, never over-reads; UNKNOWN record ids are skipped by declared length
  (forward-compatible, like unknown XML elements).
- Fixture validity is proven in the oracles: hand-built fixtures must parse in pyxlsb AND
  calamine before our tests trust them.
**Tasks**
- [x] `biff/` record stream + wide strings + RK (+ unit tests incl. the RK corner
      matrix).
- [x] `xlsb/` part parsers + `openXlsb` + seam wiring + facade export.
- [x] Fixtures: hand-built `xlsb-basic.xlsb` via the generator's binary part helpers
      (oracle-validated in pyxlsb AND calamine). A real Excel-authored workbook is a
      welcome upgrade — Excel/LibreOffice aren't available here to author one.
- [x] Tests: full cell-record matrix; RK corners (int/÷100/negative/double); style-driven
      dates; hyperlinks/visibility/dimension; truncation, lying-length, unknown-record
      degrades. *(merges + 1904 deferred — see below.)*
- [x] Adversarial review (hostile-input lens on framing; spec lens vs MS-XLSB).
**Acceptance.** Fixtures match python-calamine cell-for-cell (pyxlsb agreeing on the
hand-built ones); hostile record streams fail typed in bounded time; xlsx suite green.

**Landed (uncommitted, awaiting owner approval).** `openXlsb` returns the SAME public `Workbook`
as `openXlsx`. New `biff/record.ts` (BIFF12 record layer: variable-id + 7-bit-varint-len framing,
grounded in pyxlsb's source; RK decode int/÷100/double; UTF-16LE wide strings; EVERY field
bounds-checked so truncated/lying/malformed input degrades, never throws or over-reads). New
`xlsb/` parsers (workbook.bin → sheets + visibility; sharedStrings.bin; styles.bin → numFmtId per
cellXf, reusing the xlsx `isDateFormatCode`/`isBuiltinDateId` — the latter newly exported from
ooxml/styles) and `xlsb/sheet.ts` (the cell-record walker). `reader/xlsb.ts` = `openXlsb` +
`XlsbWorksheet implements Worksheet`, reusing the zip reader + F1.4 rels graph verbatim (same OPC
container). Reads: values (string / RK-int / real / RK-÷100 / bool / error / cached formula),
style-driven date detection, `numberFormat(ref)`, hyperlinks (via sheet rels), dimension,
visibility; everything else degrades. Fixture `xlsb-basic.xlsb` built by generate.mjs (a JS port of
a Python builder validated against pyxlsb + calamine). **DE-RISK:** a hand-built .xlsb reads
identically in pyxlsb AND calamine; date detection is validated via the MS-XLSB BrtXF iFmt@offset-2
layout + unit tests (calamine itself does NOT date-convert .xlsb, so our reader is more capable).
Gate: biome 0 / tsc 0 / **546 tests** (+18) / values cross-checked cell-for-cell vs calamine/pyxlsb
/ xlsx suite untouched-green. Runnable example added: `examples/11-other-formats.mjs` reads a `.xlsb`
and an `.ods` into the SAME Workbook API and converts the xlsb to xlsx via the bridge (all 11 examples
green). F7.4 extends this example with `.csv` + `detectSpreadsheetFormat`.
**Adversarial review (4-lens workflow → refuting verifiers; 9 candidates → 1 CONFIRMED, 8 refuted,
fixed + pinned):** the cell style field was read as a full u32, but MS-XLSB §2.5.9 packs it as
iStyleRef (24 bits) + fPhShow (1 bit) + reserved (7 bits) — an fPhShow=1 cell (CJK phonetic
workbooks) carried a corrupted style index and silently lost date detection / number formats → mask
to `& 0xffffff`; pinned by a test that fails without the mask (the committed fixture's style bytes
all have the top 8 bits clear, so only a crafted fPhShow cell exposes it). Probes stayed in the
scratchpad.
**Scope refinements (owner-notified deviations from the F7.2 scope):**
- **Merges DEFERRED for xlsb.** No independent oracle can verify a merge record without a real
  Excel-authored .xlsb (pyxlsb doesn't parse merges; calamine's `to_python` doesn't surface them),
  so shipping an unverifiable merge parser would violate the "crafted fixtures must parse in the
  oracles" rule. `mergedCells` degrades to `[]`; revisit with a real .xlsb.
- **1904 date system DEFERRED (defaults to 1900).** The `BrtWbProp` f1904 flag bit couldn't be
  verified against an oracle, and reading the WRONG bit would risk mis-dating the far more common
  1900 files — so xlsb dates use the 1900 epoch, documented, pending a real 1904 .xlsb.

### F7.3 — `.csv` / `.tsv` read: `openCsv` ☐
**Context.** Delimited text is the universal export (Google Sheets, Excel "Save as CSV",
every DB / BI / analytics tool) and the single most common "user uploaded a spreadsheet"
case. No container, no XML, no binary records — but "robust CSV" is deceptively subtle,
which is exactly why a careful zero-dep implementation earns trust: RFC 4180 quoting
(`""` escapes a quote inside a quoted field), embedded newlines / delimiters / quotes
inside quoted fields, CRLF vs LF vs CR line endings, a UTF-8 BOM, and delimiter variety
(comma, tab, and the semicolon Excel writes in some locales).
**Scope (in).** `openCsv(source, options?)` → the shared Workbook surface (one sheet — a
CSV is a single table). A hand-rolled character-scanner state machine over the decoded
text (never a regex split — that mis-handles quoted delimiters/newlines): field / quoted-
field / quote-in-quote / row transitions, CR / LF / CRLF normalized, BOM stripped. Cells
carry a positional A1 ref. **Type inference is conservative:** a field is a `number` iff
it round-trips through `Number()` as a finite value and matches a plain numeric shape
(optional sign, digits, one dot, optional `e`-exponent — NOT `Infinity`/`NaN`/hex/`0x`);
`TRUE`/`FALSE` (case-insensitive) → `boolean`; everything else stays `string`. **Dates are
NEVER inferred** (CSV date formats are locale-ambiguous — `01/02/03` — so guessing
fabricates wrong values; they stay strings, documented). An empty field → `empty`.
**Options (a real format needs knobs xlsx/ods don't — an additive, documented exception to
"one options type"):** `CsvReadOptions { delimiter?: ',' | '\t' | ';' | 'auto';
sheetName?: string; inferTypes?: boolean }`. Default `delimiter: 'auto'` sniffs the first
line (most-frequent of `, \t ;` outside quotes); `sheetName` defaults to `"Sheet1"`;
`inferTypes` defaults `true` (set false for all-strings). `.tsv` is just `delimiter: '\t'`.
**Scope (out — named).** Non-UTF-8 encodings (Latin-1/UTF-16 — documented, UTF-8 only like
the rest of the lib); header-row → object mapping (a consumer concern, not the reader's);
multi-table / multi-file; a streaming `openCsv` variant (buffered read only, like the
other new readers); quote/escape dialects beyond RFC 4180 + the `""` convention.
**Design decisions (made).**
- One `OdsWorksheet`-style plain data holder (`CsvWorksheet implements Worksheet`): cells +
  a synthesized `dimension`; every other accessor DEGRADES (no merges/styles/formula/
  comments/geometry/images) — the F7.1 pattern reused.
- Tolerant + bounded: an unterminated quote consumes to EOF as one field (bounded by input
  size — never a hang); rows past `MAX_ROW` and columns past `MAX_COL` are dropped (the
  shared grid clamp); ragged rows are fine (sparse cells). No throw for malformed content —
  only a typed error for a genuinely un-decodable source.
- No zip, so `maxPartBytes` doesn't apply; the input bytes ARE the bound. `openCsv` accepts
  `Uint8Array | ArrayBuffer | string` (text passed straight through, decoded once otherwise).
**Tasks**
- [ ] `csv/` dir: the scanner state machine → typed cell rows (quoting, embedded
      newlines/delimiters, CRLF/CR/LF, BOM, delimiter auto-sniff, conservative inference).
- [ ] `openCsv` + `CsvWorksheet` (reuse the seam) + `CsvReadOptions` + facade export.
- [ ] Fixtures: crafted `.csv`/`.tsv` via the generator (quoted fields with commas +
      embedded newlines + escaped quotes; CRLF; BOM; semicolon-delimited; a numeric/boolean/
      string type matrix; a ragged file; a date-looking column that must stay string).
- [ ] Tests: parse matrix vs Python's stdlib `csv` (RFC 4180 reference) cell-for-cell;
      delimiter sniffing; inference boundaries (`"007"` stays string? leading-zero decision
      documented; `1e3`, `-0`, `1.5`, `NaN` cases); grid-clamp bounds; the bridge
      (`workbookToInput` → `writeXlsx`) round-trips values.
- [ ] Adversarial review (hostile-input lens on quote/newline pathologies; model lens on
      inference correctness — no fabricated dates, no lost precision on big integers).
**Acceptance.** The parse matrix matches Python's `csv` cell-for-cell (quoting + embedded
newlines exact); a pathological unterminated-quote / million-column file stays bounded;
inference never fabricates a date and never corrupts a string that looks numeric-ish;
`openXlsx` behavior and the full existing suite are untouched.

### F7.4 — Detection, conversion corpus, docs + bench lanes ☐
**Scope (in).** `detectSpreadsheetFormat(bytes)` → `'xlsx' | 'xlsb' | 'ods' | 'csv' |
undefined` (zip signature → peek `[Content_Types].xml` for xlsx/xlsb vs `mimetype` for
ods; NOT a zip but decodes as text → `'csv'` best-effort; else `undefined`). CSV has no
magic bytes, so its detection is a documented heuristic (the text fallback), never as
certain as a container sniff. Cross-format equivalence corpus: ONE logical table authored
as .xlsx / .xlsb / .ods / .csv reads to the SAME value snapshot across all four readers
(modulo the documented CSV type-inference boundary — no dates); the corpus property
extends across formats — every readable non-xlsx fixture CONVERTS (`workbookToInput` →
`writeXlsx` → `openXlsx`) with values/types/merges lossless or fails TYPED. Example
`11-read-anything.mjs` (detect → open → convert to xlsx). Docs: README matrix ("reads
xlsx/xlsm, xlsb, ods, csv/tsv — writes xlsx") + per-format fidelity/drop tables in all
three READMEs (+ the CLAUDE.md repo map gains the new dirs). Verify `.xlsm`/`.xltx`
already open today (vba/template parts ignored via rels) and pin it with a fixture.
Tree-shaking check: a consumer bundling only `openXlsx` must not carry the BIFF/CSV/ODS
code (`sideEffects` honest; measured with a small esbuild probe in the scratchpad). Bench:
read lanes for xlsb/ods vs python-calamine and csv vs Python `csv`, published to
docs/benchmarks.md.
**Tasks**
- [ ] `detectSpreadsheetFormat` + tests (each format + junk + empty + truncated + a
      text-but-not-CSV edge).
- [ ] Equivalence corpus (four formats, one snapshot) + the cross-format conversion
      property in the corpus test.
- [ ] Extend example 11 (`11-other-formats.mjs` exists from F7.2 — reads .xlsb + .ods) with
      `.csv` + `detectSpreadsheetFormat`; keep examples README/package.json wiring; all green.
- [ ] README/core/facade docs matrix + per-format drop lists + CLAUDE.md repo map;
      xlsm/xltx fixture pin.
- [ ] Tree-shake probe; bench read lanes + benchmarks.md refresh.
- [ ] Full-milestone adversarial review (cross-format lens), M5/M6-analysis style.
**Acceptance.** One snapshot across four formats; the conversion property holds over
the whole corpus; docs enumerate per-format drops exactly; the bench table carries the
new lanes. Release prep + the 0.7 bump happen ONLY at the owner's explicit request
(CLAUDE.md #4).

---

## M8+ — Later milestones (outline; expanded when reached)

- **Deferred — legacy `.xls` (BIFF8) read:** a CFB/OLE2 container reader + BIFF record
  layer (globals + per-sheet substreams; the SST `Continue`-split Unicode-flag trap;
  `FILEPASS`/`VelvetSweatshop` encryption → typed reject). Cut from M7 (owner, 2026-07-10)
  as the milestone's biggest lift for a declining ~1997–2007 install base. Revisit if
  users ask; the multi-format seam (F7.1) already accommodates another `open*` + backend.
- **Deferred — native lane:** the optional `@openjsxl/native` napi-rs binding to Rust
  `calamine` (and a WASM build) behind the zip/xml interface. Deferred by the F5.5
  benchmark evidence (see the M6 re-scope note); revisit if a workload shifts the math.
- **M8 — Formula evaluation (v0.8):** opt-in parser + evaluator for common functions,
  behind a separate entry point so the core stays lean (text fidelity landed in F5.4).
- **M9 — Breadth + hardening (v0.9):** tables, data validation, conditional formatting;
  fuzzing, expanded corpus.
- **1.0:** frozen API, full round-trip fidelity, documentation site, benchmarks kept
  current (harness from F5.5).

---

## Risks & mitigations

- **OOXML edge cases.** Happy path is a weekend; Excel's real behavior is months. Mitigate
  with a golden-file corpus (Excel + LibreOffice + Google Sheets) and a ruthless won't-do
  list.
- **Writer correctness.** Excel silently repairs/rejects bad files — the reason the reader
  ships first; gate the writer on round-trip + real-app open tests.
- **Maintenance gravity.** ExcelJS shows what scale does to a solo maintainer. Mitigate with
  a GitHub org (not a personal repo), narrow scope, no paid-tier promises, CI on the corpus.
- **Native-core temptation.** Keep `calamine`/napi out of the MVP; it forfeits the
  browser/edge story and adds a prebuilt-binary release matrix before there are users.
