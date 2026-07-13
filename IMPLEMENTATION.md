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

### F7.3 — `.csv` / `.tsv` read: `openCsv` ☑
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
- [x] `csv/` dir: the scanner state machine → typed cell rows (quoting, embedded
      newlines/delimiters, CRLF/CR/LF, BOM, delimiter auto-sniff, conservative inference).
- [x] `openCsv` + `CsvWorksheet` (reuse the seam) + `CsvReadOptions` + facade export.
- [x] Fixture: crafted `basic.csv` via the generator (quoted comma, embedded newline, escaped
      quote, leading-zero ids, mixed types, CRLF). The `.tsv` / BOM / semicolon / ragged /
      date-stays-string cases are covered by inline-string tests (openCsv accepts a string).
- [x] Tests: parse matrix vs Python's stdlib `csv` (RFC 4180 reference) cell-for-cell (18 cases);
      delimiter sniffing; inference boundaries (`007`/`00` stay strings; big-int precision;
      `1e3`, `-0`, `NaN`/`Infinity`/`0x`/`1,000` stay strings); the bridge round-trips values.
- [x] Adversarial review (hostile-input on quote/newline pathologies + grid materialization;
      model lens on inference correctness).
**Acceptance.** The parse matrix matches Python's `csv` cell-for-cell (quoting + embedded
newlines exact); a pathological unterminated-quote / million-column file stays bounded;
inference never fabricates a date and never corrupts a string that looks numeric-ish;
`openXlsx` behavior and the full existing suite are untouched.

**Landed (uncommitted, awaiting owner approval).** `openCsv(source, options?)` — SYNCHRONOUS (CSV
has no container to decompress), accepting `Uint8Array | ArrayBuffer | string` and returning the
SAME public `Workbook` as `openXlsx` (one sheet). `csv/parse.ts`: an RFC 4180 char-scanner
(`parseDelimited` — quoted fields, `""` escape, embedded delimiters/newlines, CRLF/CR/LF, BOM strip;
a quote is special ONLY at field start; unterminated quote → one field to EOF, bounded) +
conservative inference (`inferCsvValue` — TRUE/FALSE→boolean; a plain numeric literal→number EXCEPT
a leading-zero integer or a big integer beyond `Number.MAX_SAFE_INTEGER`, both kept STRING; dates
NEVER inferred) + `sniffDelimiter`. `reader/csv.ts`: `openCsv` + `CsvWorksheet implements Worksheet`
(cells + synthesized dimension; everything else degrades) + `CsvReadOptions` (delimiter / sheetName /
inferTypes). The scanner is pinned CELL-FOR-CELL against Python's stdlib `csv` on 18 cases; fixture
`basic.csv` cross-checked. Gate: biome 0 / tsc 0 / **580 tests** (+34) / xlsx suite untouched-green.
**Adversarial review (3-lens workflow → refuting verifiers; 2 candidates → 0 CONFIRMED):** the
mid-field-quote divergence (a `"` mid-field was wrongly opening a quoted region — `a"b,c` →
`[["ab,c"]]` instead of `[['a"b','c']]`) was found EMPIRICALLY and fixed BEFORE the verify pass (quote
special only at `field === ""`; pinned by 4 Python-validated cases), so the verifier confirmed the
current code correct; the grid-materialization "DoS" was refuted — the intermediate `string[][]` is
O(input) with no amplification (unlike ods there's no repeat/expansion), empirically bounded
(200k-col grid ⇐ ~1.3 MB input, 26 ms / 35 MB). Probes stayed in the scratchpad.
**Note (owner-notified):** the runnable `.csv` example + `detectSpreadsheetFormat` land in F7.4
(which extends `examples/11-other-formats.mjs`); F7.3 ships the reader + tests + fixture.

### F7.4 — Detection, conversion corpus, docs + bench lanes ☑ (bench lanes → post-M7 pass)
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
- [x] `detectSpreadsheetFormat` + tests (each format + junk + empty + truncated + a
      text-but-not-CSV edge).
- [x] Equivalence corpus (four formats, one snapshot) + the cross-format conversion
      property in the corpus test.
- [x] Extend example 11 (`11-other-formats.mjs` exists from F7.2 — reads .xlsb + .ods) with
      `.csv` + `detectSpreadsheetFormat`; keep examples README/package.json wiring; all green.
- [x] README/core/facade docs matrix + per-format drop lists + CLAUDE.md repo map;
      xlsm/xltx fixture pin.
- [x] Tree-shake probe. — [ ] bench read lanes + benchmarks.md refresh (→ post-M7 bench pass).
- [x] Full-milestone adversarial review (cross-format lens), M5/M6-analysis style.
**Acceptance.** One snapshot across four formats; the conversion property holds over
the whole corpus; docs enumerate per-format drops exactly; the bench table carries the
new lanes. Release prep + the 0.7 bump happen ONLY at the owner's explicit request
(CLAUDE.md #4).

**Landed (uncommitted, awaiting owner approval).** `detectSpreadsheetFormat(source, options?)` →
`'xlsx' | 'xlsb' | 'ods' | 'csv' | undefined` (`reader/detect.ts`, facade-exported with
`SpreadsheetFormat`). Zip-container sniff reusing the hardened `openZip`: an ODF `mimetype` →
`'ods'` (a **mimetype-less ODS** is still recognized via its `content.xml` `office:spreadsheet`
body, so detect is never stricter than `openOds`); `[Content_Types].xml` binary main type → `'xlsb'`,
the four XML spreadsheet content types (.xlsx/.xlsm/.xltx/.xltm) → `'xlsx'`. Non-zip decodable UTF-8
text → `'csv'` (documented best-effort; C0-control/invalid-UTF-8 → `undefined`). Each classification
read is a **bounded 1 MiB streaming prefix** (`sniff`), so a decompression-bomb part can't force an
unbounded inflate (empirically: a 57 KB zip declaring a 60 MB Content_Types classifies in 9 ms /
~6 MB RSS). Fully type-safe — no `any`, no `as` assertions.
Cross-format **equivalence corpus** `equiv.{xlsx,xlsb,ods,csv}` (one logical table; no dates) →
identical full-cell value snapshot across all four readers (`reader/__tests__/cross-format.test.ts`;
.xlsb/.ods lanes cross-checked in python-calamine). **Conversion property** extended
(`writer/__tests__/bridge-styles.test.ts`): every readable ods/xlsb/csv fixture converts to `.xlsx`
lossless (values/types/merges/hyperlinks/state) or fails typed — the typed-reject set is pinned
(empty) and the key fixtures pinned present. **xlsm/xltx pin:** crafted `xlsm-macro.xlsm` — `openXlsx`
reads it (vba/content-type-label ignored), detect → `'xlsx'`. Example 11 rewritten as a detect-driven
"read anything" demo (11/11 examples green). Docs: format matrix + per-format drop tables in all three
READMEs; CLAUDE.md repo map (ods/biff/xlsb/csv); corpus README. **Tree-shake probe** (scratchpad
esbuild): an `openXlsx`-only bundle drops ~16.8 KB (34%) and leaks **zero** ods/xlsb/csv/detect
symbols — `sideEffects:false` honest.
Gate: biome 0 / tsc 0 / **603 tests** (+23) / 11 examples green / xlsx suite untouched.
**Adversarial review (workflow `wln823b14`: 4-lens finders → refuting verifiers; 5 candidates → 5
CONFIRMED = 2 code defects + 2 test holes, all fixed + pinned; 0 refuted):** (1) detect missed a
mimetype-less ODS `openOds` reads → added the content.xml `office:spreadsheet` fallback; (2) detect
fully inflated the mimetype/Content_Types part (a ~4.29 GB bomb DoS, contradicting its own
"never hangs") → bounded 1 MiB streaming `sniff`; (3) the conversion test didn't pin *which* fixtures
typed-reject → pinned the (empty) reject set + key-fixture presence; (4) the equivalence test sampled
9 fixed refs → now compares the full populated-cell set (catches an out-of-range cell). Probes stayed
in the scratchpad.
**Remaining for M7 (post-commit, owner's separate post-M7 ask):** the bench read lanes + library
size matrix — scoped as its own section below, its own commit.

---

## Post-M7 — bench read lanes + library size matrix ☑

The owner's post-M7 ask (2026-07-11): re-run `pnpm bench` on the M7 build and add (a) cross-format
**read** lanes and (b) a **library size matrix**. Two owner decisions locked via AskUserQuestion:
full-scale xlsb/ods lanes with **SheetJS-authored fixtures** (the same author-then-read methodology
ExcelJS→xlsx already uses), and the size matrix covering the real competitors. Private `@openjsxl/bench`
only — no shipped code changes, so the byte-identity / corpus / openpyxl gates are N/A; the standing
gate (biome/tsc/vitest) still holds since nothing in `packages/core` moves.

**Cross-format read benchmark.** One workload (`numbers`), the three existing sizes (10k/100k/1M),
four container formats, each read by the libraries that support it:
- **xlsx** — openjsxl · ExcelJS · SheetJS (reuse the ExcelJS-authored fixtures already cached).
- **xlsb** — openjsxl · SheetJS (ExcelJS/openpyxl can't read it). Fixture authored by SheetJS.
- **ods** — openjsxl · SheetJS. Fixture authored by SheetJS.
- **csv** — openjsxl · ExcelJS · SheetJS. Fixture authored directly (csv is just text — no producer).
Python reference (out-of-band, same as F5.5): **python-calamine** reads xlsx/xlsb/ods (the native
speed bar); **Python stdlib `csv`** for csv. Equal-work rule unchanged: every reader materializes
every cell into a checksum sink, and all readers of a format parse the identical file. **Caveat noted
in the report:** SheetJS reads xlsb/ods files it also authored (the same self-authored shape ExcelJS
already has for xlsx) — the numbers are throughput-indicative, and calamine (independent) anchors them.

**Library size matrix.** openjsxl (facade) · @openjsxl/core · ExcelJS `4.4.0` · SheetJS `xlsx@0.18.5`,
by three honest measures: **runtime dependency count** (direct + transitive), **installed footprint**
(the package + its deps on disk), and **minified+gzipped bundle size** (esbuild bundle of the main
entry — the browser/edge cost, where openjsxl's zero-dep story shows). Reproducible via a small
`sizes.mjs` in the bench package; every published number comes from a measurement, never an estimate.

**Tasks**
- [x] Adapters: `read(bytes, format)` dispatch (openjsxl → open{Xlsx,Xlsb,Ods,Csv}; SheetJS auto-detects;
      ExcelJS gains csv, declines xlsb/ods); worker passes `format`.
- [x] `formats.mjs`: the per-format reader matrix + SheetJS/direct fixture authoring (cached).
- [x] `run.mjs`: a format-read phase after the main matrix; `report.mjs`: a "Read by format" section.
- [x] `bench_py.py`: calamine xlsb/ods read lanes + a Python stdlib `csv` lane; merge into the report.
- [x] `sizes.mjs` + a "Library size" section in benchmarks.md; a size line in the three READMEs.
- [x] Run it (JS lanes + a fresh calamine venv for the Python reference); refresh docs/benchmarks.md.
**Acceptance.** benchmarks.md carries the cross-format read table + the size matrix, every number from
a real run; READMEs cite the size comparison; the existing xlsx read/write numbers are refreshed on the
M7 build. Own commit; 0.7 release prep + bump stay owner-gated (CLAUDE.md #4).

**Landed (uncommitted, awaiting owner approval).** Adapters gained `read(bytes, format)` dispatch
(openjsxl → the four typed openers; SheetJS auto-detects; ExcelJS reads csv, declines xlsb/ods typed);
`worker.mjs` threads `format`. New `formats.mjs` (per-format reader matrix; `.xlsb`/`.ods` fixtures
authored by SheetJS, `.csv` written directly — cached under `.cache/`); `run.mjs` runs a format-read
phase after the main matrix; `report.mjs` renders **Library size** + **Read by format** sections and a
Python by-format table. New `sizes.mjs` (`pnpm sizes` → clean `npm install --omit=dev` per lib →
`.cache/sizes.json`). `bench_py.py` gained calamine xlsx/xlsb + stdlib-`csv` format lanes.
**Numbers (M2 Pro, Node 24, 2026-07-11):** size — openjsxl **0 third-party deps / 0.2 MB installed**
vs ExcelJS 97 pkgs / 34 MB vs SheetJS 9 pkgs / 14 MB. Read-by-format — openjsxl leads every format at
every size; **`.xlsb` 1M in 0.18 s vs SheetJS 1.55 s (~8×)**, `.ods` 1.39 s vs 3.30 s, `.csv` 0.59 s
vs 1.04 s. **Finding surfaced (not hidden):** python-calamine **rejects SheetJS's `.ods` output** with
a parse error while openjsxl and SheetJS both read it — so the `.ods` lane has no calamine anchor (an
openjsxl-tolerance data point). Docs: docs/benchmarks.md (size + format + Python-by-format sections),
root README (refreshed 1M table + size mini-table + cross-format line), facade README (size line), bench
README (`pnpm sizes` + the new lanes). Gate: biome 0 / tsc 0 / **603 tests** (no core code moved — bench
is a private package). **M7 fully complete.**

**Post-M7 review follow-up (uncommitted, awaiting owner approval).** Owner-requested cost-capped
review (workflow `wmb7424k9`: 3 finders → capped refuting verifiers, 7 agents) of `5cc8802` + `c2f6f6a`
— targeting the F7.4 post-review fixes (themselves unreviewed) and the never-reviewed bench commit.
**7 findings: 4 CONFIRMED by verifiers + 3 UNVERIFIED (cost cap) re-verified CONFIRMED by hand; 0
refuted. All fixed:**
1. *(medium, core)* `detect.ts`: the content.xml ODS fallback fired even when a present non-empty
   NON-spreadsheet `mimetype` had been read — classifying files `openOds` deterministically rejects
   (a real `.odt`, a garbage-mimetype zip) as `ods`. Fixed: the fallback now fires only when the
   mimetype entry is absent or empty (mirroring ods.ts's own tolerance) and matches the structural
   `<office:spreadsheet` element, not the bare substring. +3 pinned tests (**606**).
2. *(low, core)* The "detect is never stricter than its opener" comment/invariant was false for a
   mimetype-less `.odt` (openOds opens it as a ZERO-sheet workbook; detect says undefined —
   deliberately). Comments in detect.ts/detect.test.ts rewritten honestly; openOds's docstring
   corrected (non-spreadsheet rejection happens via the mimetype, when present).
3. *(medium, bench)* `sizes.mjs` counted the library itself as its own dependency (off-by-one:
   openjsxl "2"→**1** (its own core), ExcelJS 97→**96**, SheetJS 9→**8**); report.mjs's bold-0 branch
   was unreachable. Fixed count (`lines − 2`), sizes.json regenerated, READMEs reconciled.
4. *(low, bench)* The sizes.json `note` (openjsxl measured at published 0.6.0; M7 adds no deps) was
   captured but never rendered — now a footnote under the size table.
5. *(low, bench)* formats.mjs comment claimed calamine anchors the **ods** lane — false (calamine
   rejects SheetJS-authored ods); comment now states the ods lane has no independent anchor.
6. *(medium, bench)* The Python reference ran on **calamine 0.7.0 / Python 3.14.6** while
   requirements.txt pinned 0.2.3 / ≤3.13 — published numbers weren't reproducible from the documented
   env. requirements.txt re-pinned to the versions actually measured (0.7.0 ships 3.9–3.14 wheels);
   bench README updated.
7. *(low, docs)* "within ~1.5× of native calamine" was stale vs the fresh data (1M read ratios
   1.31/1.61/1.63) → reworded to "~1.3–1.6× (workload-dependent)".
Plus an owner style pass on detect.ts: the blanket try/catch around the whole zip branch replaced by
per-site shields on the only two throwing ops — one try around `openZip`, and `sniff()` owning the
try around its readStream iteration (→ `string | undefined`); pure logic stays shield-free. Side
benefit, pinned: one unreadable entry (e.g. an oversized mimetype under `maxPartBytes`) no longer
aborts classification — the OOXML branch still decides (openXlsx never reads mimetype).
Gate: biome 0 / tsc 0 / **607 tests** (+4). Probes stayed in the scratchpad; repo swept clean.

---

## M8 — Formulas (v0.8)

**Theme.** An **opt-in** formula parser + evaluator behind a **separate entry point** —
`openjsxl/formula` — so the core stays lean: a consumer who never imports it never loads a
byte of it (module-graph isolation, stronger than tree-shaking). Formula TEXT
read/translate/write shipped in F5.4 (`ooxml/formula.ts` is a shift tokenizer for
shared-formula translation, NOT an AST — F8.1 builds the real parser). Scoped 2026-07-11
from a 2-agent research pass with install-verified oracles.

**Milestone-wide decisions** (set here so features don't re-litigate):
1. **Separate subpath entry, not a new package.** `@openjsxl/core` gains `"./formula"`
   (tsup `entry: { index, formula: 'src/formula/index.ts' }`; `exports["./formula"]` with
   its own `types`/`import`), and the `openjsxl` facade mirrors it (`src/formula.ts` =
   `export * from "@openjsxl/core/formula"` — the facade header already reserves this
   pattern). Same-package placement lets the evaluator reuse `a1.ts`/`dates.ts` internals
   without exposing them as public API. **Build rule:** tsup gets `splitting: false` —
   a second ESM entry otherwise hoists shared modules into chunks and would churn
   `dist/index.js`; with splitting off, the formula bundle duplicates the few shared
   helpers (a few KB — accepted) and the `"."` entry's emitted file stays byte-identical.
   Verify exactly that, plus `writeXlsx` bytes and a tree-shake/size probe re-run (F7.4
   precedent).
2. **Evaluation is read-only.** `evaluate*` never mutates the `Workbook`, never writes
   results into cells, and the writer never refreshes stored `<v>` caches (byte-identity).
   Our evaluated result CAN disagree with a stale producer cache — expected, documented.
3. **Deterministic by default.** Volatile functions (NOW, TODAY, RAND, RANDBETWEEN)
   **typed-reject** unless the caller injects `options.now` / `options.random` — a
   configuration failure naming the cell + function, distinct from in-sheet error values.
   No fixed-epoch default: a silently-wrong date is the silent mangling this repo bans.
4. **Errors are values; failures are typed.** The full 8-member ST_CellErrorType set —
   `#DIV/0! #VALUE! #REF! #NAME? #N/A #NUM! #NULL! #GETTING_DATA` — parses and PROPAGATES
   as `EvalValue` variants that round-trip into the writer's error-cell path
   (`#GETTING_DATA` already flows through the shared Cell model via the xlsb reader; our
   own functions never PRODUCE it). Cycles evaluate to a
   dedicated cycle-error value per affected cell (HyperFormula precedent; NOT Excel's
   silent 0, and NOT a throw — one cycle must not abort an unrelated region). Parser/
   evaluator failures (depth cap, fuel budget, volatile-unconfigured) throw a typed
   `FormulaError` — the formula entry's OWN error class; core's closed `XlsxErrorCode`
   union stays frozen.
5. **Scalar-only evaluation; every exclusion named, parseable, never silent.** Excluded
   from v0.8 evaluation (still PARSED so nothing is mangled): structured refs
   `Table1[@Col]`, external refs `[1]!A1` (→ typed unsupported error value), dynamic-array
   spill `#` (array masters yield their top-left value — named degradation), R1C1
   (excluded outright — never the stored form). `@` implicit intersection parses AND
   evaluates (scalar passthrough / single-cell intersection / else `#VALUE!`). 3-D refs
   (`Sheet1:Sheet3!A1`) are **parse-only in v0.8** — evaluation yields the typed
   unsupported error value (a multi-sheet RangeView + sheet-order resolution is its own
   work item; promising "aggregators only" without owning it somewhere was scope debt).
   Defined names resolve constants + simple ranges only; anything else →
   `#NAME?`/typed-unsupported (scope fence — `refersTo` is a full formula language of its
   own). At 418 functions, HyperFormula ships without structured refs/3-D/dynamic arrays
   — good cover.
6. **Excel semantics, plain IEEE-754 doubles.** The verified coercion matrix is the
   contract: empty cell → 0 in arithmetic, `""` in concat, equals BOTH 0 and `""` under
   `=`; TRUE→1/FALSE→0; numeric strings coerce in arithmetic (locale-invariant en-US
   parsing, never host-locale) else `#VALUE!`; comparisons never coerce and order
   number < text < FALSE < TRUE (text case-insensitive); SUM skips text/bools in RANGES
   but coerces literal ARGS (`SUM("5",TRUE)`=6); IF/CHOOSE evaluate only the taken branch,
   AND/OR do NOT short-circuit; `=-2^2`→4 (unary minus binds above `^`), `=2^3^2`→64
   (left-assoc `^`). No cosmetic display rounding (no evaluator replicates it — pycel,
   'formulas', HyperFormula are all plain doubles ± epsilon); oracle comparisons use
   relative tolerance (~1e-9).
7. **Adversarial-input safety by construction.** The stored form allows a ~4,000-level
   paren bomb inside MAX_FORMULA_LEN=8192 (the 64-level limit is FUNCTIONS only) — the
   parser uses an explicit stack/depth caps (64 function nesting, 256 parens, ≤255 args)
   with typed errors, never a bare RangeError. Ranges NEVER materialize: `SUM(A:A)`
   iterates USED cells via a lazy `RangeView` (F4.4/F4.6 precedent); `COUNTBLANK` does
   extent arithmetic. Evaluation carries a fuel budget (`options.maxCellVisits`, generous
   default) → typed budget-exceeded; memoization makes shared subgraphs pay once. The
   evaluator itself is iterative (explicit work stack + tri-color marks) — a 1M-row
   dependency CHAIN is acyclic and must not blow the JS stack.
8. **Oracles (install-verified 2026-07-11).** PRIMARY: Python **`formulas`** 1.3.4 (EUPL,
   venv-only) — works on Py 3.14, evaluates raw strings AND whole .xlsx workbooks, 470
   functions. SECONDARY: **pycel** 1.0b30 (GPLv3, venv-only) — BROKEN on Py ≥3.12
   (`ast.Str`), pin `/opt/homebrew/bin/python3.11` explicitly. TERTIARY: Excel-authored
   fixtures' own cached `<v>` values (Excel computed them; zero tooling). **HyperFormula
   as an in-repo devDep is an OWNER DECISION** (GPLv3-or-commercial dual license +
   `licenseKey:'gpl-v3'` in test code) — default is scratchpad-npm probes only, the
   openpyxl-venv pattern. formulajs is MIT but parser-less (function-impl reference only).

**Standing gates** per feature: biome by exit code + tsc + vitest; adversarial review with
confirmed findings fixed + pinned pre-commit. The writer is untouched all milestone —
golden pins stay green; entry-point work (F8.1) additionally re-runs the tree-shake probe
and verifies `"."` chunk-split neutrality.

**Dependency order:** F8.1 parser → F8.2 evaluator core → F8.3 function library → F8.4
integration + docs + oracle corpus. One feature per session, each proceed-gated.

### F8.1 — Formula parser + the `openjsxl/formula` entry point ☑
**Context.** The stored form is canonical en-US (ECMA-376 §18.17): `,` arg separators,
`.` decimals, quoted sheet names with `''` doubling, TRUE/FALSE, array constants
`{1,2;3,4}` (`,` = columns, `;` = rows, constants only). Full MS precedence table:
`: (space) ,` reference ops → unary `-` → `%` → `^` → `* /` → `+ -` → `&` → comparisons;
left-to-right on ties, plus the two quirks in decision 6.
**Scope (in).** `src/formula/` in core: a lexer (extends F5.4's token classes with
operators/braces/array literals) + a Pratt / precedence-climbing parser producing a typed
AST — explicit operand/operator stack or hard depth counter (decision 7 caps). Parses the
COMPLETE stored-form surface: literals (number/string/bool/all 8 error literals incl.
`#NULL!`), $-pinned A1 refs, whole-column/row ranges (`A:A`, `1:1`), quoted+unquoted sheet
refs, 3-D refs, union `,` inside parens, range `:` as operator (parse-only for
`INDEX():INDEX()`), defined names, calls ≤255 args, array literals, structured/external
refs → opaque nodes, `@`/`#` tokens. `parseFormula(text): FormulaAst` + typed
`FormulaError('parse-error' | 'depth-exceeded' | …)`. The `"./formula"` subpath lands here
on both packages.
**Scope (out — named).** No evaluation (F8.2); no R1C1; no locale variants (stored form
only); no serializer back to text beyond what translateFormula already does.
**Tasks**
- [x] Lexer + AST types + Pratt parser (explicit stack; depth/arg caps; typed errors).
- [x] `"./formula"` subpath: core tsup second entry + exports map; facade mirror; the
      `"."` chunk-split neutrality check + tree-shake probe re-run.
- [x] Tests: precedence table pinned (incl. `-2^2`, `2^3^2`, `%`, `&`), every literal
      form, sheet-name quoting, array constants, the 8,192-char paren bomb → typed error,
      255-arg cap, opaque structured/external nodes; parse–reprint round-trips where the
      grammar is unambiguous.
- [x] Oracle: token/shape agreement vs Python `formulas` `Parser().ast()` on a matrix of
      real-world formulas (incl. every F5.4 fixture formula).
- [x] Adversarial review (parser lens).
**Acceptance.** Every formula in every corpus fixture parses to an AST (or a typed error);
the paren bomb is a typed error in bounded time; the `"."` entry's bytes/behavior are
untouched (size probe re-run); core stays zero-dep.

**Landed (F8.1).** `src/formula/` = `lexer.ts` (hand-rolled scanner: cell/name/function
disambiguation, `$`-partial refs, error literals, balanced brackets) + `parser.ts`
(precedence-climbing; `MAX_DEPTH=256` recursion guard, `MAX_FUNCTION_DEPTH=64`,
`MAX_ARGS=255`) + `ast.ts` (typed `FormulaAst`) + `errors.ts` (`FormulaError`, own code union
`parse-error|depth-exceeded|too-many-args|unsupported` — core's `XlsxErrorCode` stays frozen)
+ `index.ts`. `openjsxl/formula` wired on both packages (tsup 2nd entry + `exports` map,
`splitting:false`). **Byte-identity:** `dist/index.js` = `e678d0db…` UNCHANGED; zero parser
symbols leak into it; formula bundle imports nothing external (16.4 KB); core `dependencies`
still `{}`. **Oracle:** 92/92 real-world formulas agree with Python `formulas` 1.3.4 (Py3.14) —
it surfaced the absolute whole-col/row-endpoint gap (`$A`/`$2`), fixed. **Review (parser lens,
3 finders):** hostile-input = clean (200k random + curated inputs: every failure typed, all
paths linear, nothing materialized, stack bounded). 2 CONFIRMED defects FIXED + PINNED —
(1) spill `#` on a range's right endpoint wrapped the whole range (`A1:A5#` → `(A1:A5)#`);
`SPILL_BP` was below the `:` range op → raised above it, so `#` binds to each endpoint;
(2) deleted-sheet form `#REF!!A1` (Excel's stored form when a referenced sheet is deleted) was
rejected while the deleted-cell form `Sheet1!#REF!` parsed → `#REF!` now accepted as a sheet
name. Plus 1 robustness fix: quoted external detection widened to `[ ] / \` (all illegal in a
real sheet name) so a drive/URL path (`https://…`) is not mis-split as a 3-D span.
**Conscious scope boundary (owner call):** modern dynamic-array error literals `#SPILL!`,
`#CALC!`, `#FIELD!` etc. are rejected — decision 4 froze the error set to the 8 `ST_CellErrorType`
members; not widened here. Gate: biome 0 / tsc 0 / **716 tests** (+ lexer + parser suites).
Probes stayed in scratchpad; repo swept clean.

### F8.2 — Evaluator core: dependency walk, coercion, errors, budgets ☑
**Context.** One-shot, pull-based evaluation — not a reactive engine. Pull + memoization
+ tri-color marks gives topological order implicitly, cycle detection for free (grey hit),
and sparse demand (`evaluateCell` and `evaluateAll` both fall out).
**Scope (in).** `evaluateWorkbook(workbook, options?)` / `evaluateCell(workbook, sheet,
ref, options?)` over the SHARED `Workbook` surface (any format that carries formula text —
xlsx today). Iterative walker (explicit stack, tri-color: white/grey/black; grey-hit =
cycle → cycle-error value + the cycle's refs reported in the result). `EvalValue` =
`number | string | boolean | null(empty) | FormulaErrorValue | RangeView` (lazy iterator
over USED cells — never a materialized 2-D array). The decision-6 coercion matrix; error
propagation; per-function laziness flags (IF-family lazy, AND/OR strict). Defined-names
resolution (workbook.xml `<definedNames>` incl. `_xlnm.*` + sheet-scoped `localSheetId` —
NEW parser, nothing in-tree reads it today): constants + simple ranges only per decision 5.
Fuel budget + `options.now`/`options.random` plumbing (rejection lives here; the volatile
functions themselves land in F8.3).
**Scope (out — named).** INDIRECT/OFFSET-style dynamic references (they make the dep graph
unknowable at parse time and reopen the cycle/giant-range analyses — the concrete reason
they are absent from F8.3's tiers too); iterative-calculation mode; cross-workbook refs.
**Tasks**
- [x] `<definedNames>` parser (constants + simple ranges; the rest → `#NAME?`-on-use).
- [x] Tri-color iterative walker + memo table + cycle-error values + fuel budget.
- [x] Coercion/error semantics module with the decision-6 matrix as a pinned test table
      (each trap = one oracle-validated case).
- [x] `RangeView` over the reader's sparse cells; whole-column/row intersection with the
      used range; `COUNTBLANK` extent arithmetic.
- [x] Hostile tests: self-ref `=A1` in A1, long cycle, 1M-row acyclic chain (no stack
      growth), `SUM(A1:XFD1048576)` in bounded time, budget-exceeded typed.
- [x] Adversarial review (semantics + hostile-input lenses).
**Acceptance.** The matrix table matches Python `formulas` (tolerance 1e-9) cell-for-cell;
cycles yield error values without aborting unrelated cells; all hostile cases bounded.

**Landed (F8.2).** `src/formula/`: `value.ts` (`EvalValue` = `number|string|boolean|null|
FormulaErrorValue|RangeView`; interned error values incl. eval-only `#CYCLE!`; lazy `RangeView`
— `width`/`height`/`cellCount` by arithmetic, `entries`/`values` over USED cells only,
`single`/`topLeft`/`populatedCount`), `coerce.ts` (decision-6 matrix — empty-cell vs empty-string,
TRUE→1, numeric-string, no-coerce comparisons ordered number<text<FALSE<TRUE, `^`/`/0`/`%`,
error propagation), `functions.ts` (`FunctionSpec` eager/lazy union + `EvalContext` + registry;
caller `options.functions` normalized/validated, output sanitized; built-ins empty until F8.3),
`eval.ts` (the walker + defined-names + dispatch). `evaluateWorkbook`/`evaluateCell` are async
(snapshot each sheet via `rows()`, then a SYNC walk). Reader: `parseWorkbook` now also parses
`<definedNames>`; `Workbook.definedNames: readonly DefinedName[]` added (public, additive — the
only reader-surface growth; ods/xlsb/csv default `[]`). **Engine:** the cell walk is a
generator-driven iterative driver (`runFormula`) — each cell's formula is a generator that YIELDS
its cell deps; direct-reference chains resolve via an explicit frame stack, so a 300k–1M-row
`=A1+1` chain runs in O(1) native stack (verified 300k → 1.8s). Tri-color grey-hit → `#CYCLE!`
value; unrelated cells unaffected; **order-independent** (evaluateCell == evaluateWorkbook per
cell, verified). **Oracle:** 36/36 real-workbook cells matched Python `formulas` 1.3.4 (+openpyxl,
Py3.14). Volatile gate throws typed `volatile-unconfigured` without `options.now/random`; fuel
budget → `budget-exceeded`. **Adversarial review (3 lenses; owner stopped the agents mid-run, so
completed by hand from their two leads + the oracle):** 2 CONFIRMED adversarial-safety defects
FIXED + PINNED — (1) a self-referential range formula (`A1=A1:A1`) HUNG (range construction is
lazy so cycle detection missed it) → bounded unwrap in `reduce`/`scalarize` → typed `#CYCLE!`;
(2) a deep left-associative chain (`1+1+…`, ~4000 terms/8KB — the parser builds it iteratively so
it escaped the parser's depth cap) overflowed the native stack with an UNTYPED `RangeError` →
threaded a single `depth` through the whole sync recursion (AST descent + lazy/range nesting; the
iterative frame stack resets per frame so long direct chains still work), capped `MAX_NATIVE=256`
→ typed `depth-exceeded`; verified no untyped throw across flat/paren/lazy-nest/range-nest
frame-heavy paths. **Known v0.8 edges (documented):** built-ins are F8.3 (tests use caller specs);
1904-epoch dates not surfaced; General number→text is 15-sig-digit; scalar-only (union/3-D/
structured/external → typed error value); >256-deep expression nesting typed-rejects. Gate: biome
0 / tsc 0 / **745 tests** (coerce + eval suites). Probes stayed in scratchpad; repo swept clean.

### F8.3 — Function library: tier 1 + tier 2 + the extension registry ☑
**Context.** Triple-oracle coverage exists for every tier-1 function (pycel ~164 ∩
`formulas` 470 ∩ HyperFormula 418).
**Scope (in).** **Tier 1 (the v0.8 cut, ~40):** SUM AVERAGE COUNT COUNTA COUNTBLANK MIN
MAX IF IFERROR AND OR NOT VLOOKUP HLOOKUP INDEX MATCH CHOOSE SUMIF COUNTIF AVERAGEIF ROUND
ROUNDUP ROUNDDOWN INT ABS MOD CONCAT/CONCATENATE LEN LEFT RIGHT MID TRIM UPPER LOWER
ISBLANK ISNUMBER ISTEXT ISERROR ISNA TODAY NOW (the last two behind the injection gate).
**Tier 2 (stretch, in-order, cut allowed at review):** SUMIFS COUNTIFS AVERAGEIFS
SUMPRODUCT XLOOKUP IFS SWITCH XOR SUBSTITUTE REPLACE FIND SEARCH VALUE TEXTJOIN REPT EXACT
PROPER CHAR CODE POWER SQRT EXP LN LOG LOG10 PI SIGN TRUNC CEILING FLOOR ROW COLUMN ROWS
COLUMNS MEDIAN LARGE SMALL DATE YEAR MONTH DAY HOUR MINUTE SECOND TIME DATEVALUE EDATE
EOMONTH WEEKDAY DAYS N T NA ERROR.TYPE RAND RANDBETWEEN (the last two behind the
`options.random` injection gate — they are what makes that plumbing real). **Registry:** `options.functions:
Record<string, FunctionSpec>`; `FunctionSpec = { minArgs, maxArgs, volatile?, lazyArgs?,
evaluate(args, ctx): EvalValue }`; case-insensitive; built-ins are pre-registered specs
(the registry is proven by being our own mechanism); unknown function → `#NAME?` value
(parse always succeeds). Caller-supplied specs validated with `isPlainRecord` + single-read
TOCTOU like every other caller object.
**Scope (out — named, with reasons).** TEXT (drags in a full number-format RENDERER — its
own feature); OFFSET/INDIRECT (F8.2 decision); financial/engineering/statistical breadth;
array-arithmetic inside aggregator args (`SUM(A1:A3*B1:B3)` — CSE masters evaluate
scalar-only → top-left value; NAMED degradation in the drop list).
**Tasks**
- [x] Registry + spec validation + volatile gate; tier-1 implementations w/ per-function
      oracle tests (range-vs-literal coercion asymmetry pinned per aggregate).
- [x] Tier-2 in order (each lands with its oracle test or not at all).
- [x] Lookup semantics pinned: VLOOKUP/MATCH approximate-match (sorted assumption),
      exact-match miss → `#N/A`; INDEX 0/out-of-range → `#REF!`.
- [~] UDF example + docs (registry is public API) — the registry is already public + JSDoc'd
      since F8.2; the runnable `12-formulas.mjs` example + coverage/docs table are folded into
      **F8.4** (integration), matching how F8.1/F8.2 deferred their example.
- [x] Adversarial review (function-semantics lens vs oracles) — two passes (see below).
**Acceptance.** Every shipped function agrees with documented Excel behavior on a shared
vector table (Python `formulas` used as a secondary oracle — see the KNOWN divergences
below where it is not Excel-faithful); the volatile gate typed-rejects without injection and
is deterministic with it; UDFs register and evaluate.

**Landed (F8.3, uncommitted — awaiting owner approval).** `src/formula/builtins.ts` = ~85
built-in functions (all of tier 1 + most of tier 2), typed (no `any`/`as`), returning
`EvalValue` directly (trusted — they skip the caller-spec sanitize wrapper). Wiring changes:
`RegisteredFunction` became a **discriminated union** (`EagerRegistered | LazyRegistered`) so
`evalCall` and the built-ins dispatch cast-free; `buildRegistry` seeds from `builtins.ts`'s
`BUILTIN_ENTRIES`; `coerce.numericStringValue` extracted; `RangeView.cellAt(rowOff,colOff)`
added for positional access (SUMIF/VLOOKUP/INDEX/MATCH). Aggregates apply Excel's range rule
(a `RangeView` arg counts numbers only; text/bool/blank ignored, errors propagate) vs the
literal rule for scalar args.
- **Coverage.** Math/agg: SUM AVERAGE COUNT COUNTA COUNTBLANK MIN MAX MEDIAN LARGE SMALL
  ROUND(UP/DOWN) INT ABS SIGN TRUNC MOD POWER SQRT EXP LN LOG LOG10 PI CEILING FLOOR
  SUMPRODUCT. Logical: IF IFERROR IFNA IFS SWITCH AND OR XOR NOT. Lookup: VLOOKUP HLOOKUP
  MATCH INDEX CHOOSE ROWS COLUMNS. Cond-agg: SUMIF COUNTIF AVERAGEIF SUMIFS COUNTIFS
  AVERAGEIFS. Text: CONCAT CONCATENATE TEXTJOIN LEN LEFT RIGHT MID TRIM UPPER LOWER PROPER
  REPT EXACT CHAR CODE VALUE SUBSTITUTE REPLACE FIND SEARCH. Info: ISBLANK ISNUMBER ISTEXT
  ISLOGICAL ISERROR ISERR ISNA N T NA ERROR.TYPE. Date: DATE YEAR MONTH DAY HOUR MINUTE
  SECOND WEEKDAY TIME DAYS EDATE EOMONTH. Volatile (gated): TODAY NOW RAND RANDBETWEEN.
- **Deferred (named, documented).** ROW/COLUMN (need reference-position plumbing — a ref arg
  collapses to a scalar; like OFFSET/INDIRECT), XLOOKUP, DATEVALUE, TEXT (number-format
  renderer), and wildcard **SEARCH** (a substring glob is O(n²) over attacker-controlled cell
  text — SEARCH matches `*`/`?` literally in v0.8).
- **Known non-Excel-faithfulness of the oracle** (we follow Excel, not `formulas`): the oracle
  coerces numeric text INSIDE a reference for SUM/COUNT/AVERAGE (Excel ignores non-numbers in a
  range); keeps TRIM's internal space runs (Excel collapses them); rejects VALUE("50%") and
  mis-handles COUNTIF("?"). **Documented degradation** (F8.2 arg contract): a single-cell ref
  collapses to a scalar, so an aggregate treats it as a literal (text/bool in a single
  referenced cell coerces rather than being ignored); multi-cell ranges are Excel-exact.
  ROUND uses naive half-away (Excel's binary-fudge on e.g. 1.005 not replicated — matches the
  oracle); DATE pre-1900-03-01 inherits `dates.ts`'s 1900-leap-bug off-by-one. *IFS where
  EVERY criterion matches blank (rare) may under-count blank positions (bounded-perf tradeoff).
- **Adversarial review — pass 1** (4 lenses × refuting verifiers, oracle-wired): 22 findings,
  **18 CONFIRMED**, all fixed + regression-pinned. 11 distinct bugs: *IF family skipped
  blank cells that satisfy `"<>x"`/`""` (now a bounded blank pass + blank-excluding driver
  selection); CEILING/FLOOR rejected (neg number, pos significance); SUMPRODUCT didn't
  propagate errors where the first factor was 0/blank (now a bounded pre-scan); COUNT
  propagated errors (now ignores them, Excel-correct); WEEKDAY lacked types 11–17; **ReDoS**
  in wildcard translation (regex `.*a.*a…` was exponential → replaced with a linear
  two-pointer glob matcher); EXP overflow leaked `Infinity`; POWER(0,−) gave `#NUM!` not
  `#DIV/0!`; TRUNC huge-digits / date-serial overflow leaked `NaN`; VALUE ignored
  percent/thousands/currency; DATE pre-1900 off-by-one (documented, not fixed).
- **Adversarial review — pass 2** (focused on the pass-1 fixes: the *IF blank/driver logic,
  the glob matcher, the numeric/date guards): **2 CONFIRMED regressions I had introduced**,
  both fixed + pinned — the blank second pass over-counted when the value range was LARGER
  than the criteria range (now clips to the criteria rectangle, so Excel's reshape holds), and
  `toSerial` rejected a time-of-day on 9999-12-31 (now compares the truncated day).
**Gate:** biome 0 / tsc 0 / **775 tests** (30 new in `builtins.test.ts`). Byte-identity of the
`"."` entry unaffected (formula code is isolated behind the `openjsxl/formula` subpath; the
changeset touches only `formula/` + docs).

### F8.4 — Integration, docs, example, oracle corpus ☑
**Scope (in).** End-to-end: real openpyxl/Excel-authored fixture workbooks evaluated
against (a) Python `formulas` whole-workbook results and (b) their own stored `<v>` caches
(with the stale-cache caveat documented); evaluation-vs-cache divergence NAMED. Example
`12-formulas.mjs` (parse → evaluate → UDF → error/cycle handling). README/core/facade
docs: the entry point, determinism/volatile contract, the drop list (decision 5 + F8.3
outs), function coverage table. Bench lane: evaluate-1M-SUM-chains vs HyperFormula
(scratchpad probe) — indicative only. Size matrix re-run (dist gains the formula chunk;
`"."` unchanged). Full-milestone adversarial review, M7-style.
**Acceptance.** Corpus formulas evaluate to oracle-agreed values or NAMED degradations;
docs enumerate the exclusions exactly; `pnpm check`/tsc/vitest green; core `"."` bytes +
writeXlsx bytes untouched; 0.8 release prep owner-gated as always.
**Tasks**
- [x] Integration/oracle-corpus tests (`formula/__tests__/integration.test.ts`): a realistic
      multi-function workbook evaluated whole (12 cells, oracle-cross-checked out-of-tree);
      cross-sheet; evaluation-vs-cache AGREEMENT on the real `basic.xlsx` + `shared-formula.xlsx`
      producer fixtures; the stale-cache supersession contract (decision 2).
- [x] Example `12-formulas.mjs` (parseFormula AST → evaluateWorkbook → evaluateCell → a UDF →
      error propagation → `#CYCLE!` → volatile gate + injection) + examples README/package wiring.
- [x] Docs: root README `## Formulas (0.8)` (entry point, read-only/stale-cache, cycles, UDFs,
      determinism/volatile gate, 90+ coverage list, exact drop list) + status-line note; facade
      README paragraph; core README `## Exports` formula bullet. Coverage list verified against the
      registry (97 fns); every drop-list item verified absent (→ `#NAME?`).
- [x] Bench probe (indicative, scratchpad): 100k-deep ref chain ~0.76s, `SUM` over 100k ~0.37s
      (incl. write+read). HyperFormula head-to-head DEFERRED pending the owner's GPLv3 dev-oracle
      call. Size: core `dist/formula.js` ≈ 86 KB (the formula chunk); core `dist/index.js` (the
      `"."` entry) ≈ 195 KB, unchanged by F8.x (formula code is subpath-isolated).
- [x] Full-milestone adversarial review (M7-style): 3 lenses (holistic-semantics, cross-feature,
      docs-accuracy) + refuting verifiers → 7 findings, **3 CONFIRMED**, all fixed + pinned:
      (1) IFERROR/IFNA misread a MULTI-cell range arg as an error — `scalarize` reduces a range to a
      `#VALUE!` sentinel, so `SUM(IFERROR(A1:A3,0))` returned 0 not 6 (now the range passes through);
      (2) **3-D references** (`Sheet1:Sheet3!A1`) silently evaluated against only the first sheet —
      the parser carried `SheetSpec.toName` but eval dropped it → now a typed `#REF!` (decision 5,
      never silent); (3) README drop-list misstated `SUM(A1:A3*B1:B3)` as "→ top-left" when the
      engine returns a typed `#VALUE!` (doc corrected). Cross-feature bugs the per-feature passes
      couldn't see (they surface only when F8.2 ranges compose with F8.3 IFERROR / the parser's 3-D
      AST meets the evaluator).

**Landed (F8.4, uncommitted — awaiting owner approval).** Changeset = the two new files above + the
5 docs/example-wiring files + the two milestone-review source fixes (`formula/builtins.ts` IFERROR/
IFNA range pass-through, `formula/eval.ts` 3-D `toName` → `#REF!`) and their regression tests. All
changes are within the `formula/` subpath, so the `"."` entry and writer bytes stay untouched.
Oracle corpus: the committed integration test pins a 12-cell realistic workbook (all 12 agree with
Python `formulas` out-of-tree) AND asserts our re-evaluation equals openpyxl's cached `<v>` on two
real producer fixtures — the cross-engine corpus check, in-tree and reproducible. **Gate:** biome 0
/ tsc 0 / **782 tests**. **M8 (formulas) is feature-complete → 0.8 release prep is owner-gated.**

---

## M9 — Breadth + hardening (v0.9)

**Theme.** The three highest-demand structural features after styles — **tables, data
validation, conditional formatting** — read + write + bridge, plus a **fuzzing harness**
and corpus expansion that harden every reader against wild bytes. Scoped 2026-07-11 from
schema + probe research (element order verified against ECMA-376 sml.xsd AND openpyxl
output AND the in-tree `inventory-table.xlsx`).

**Milestone-wide decisions:**
1. **Element order (load-bearing, exact slots).** In CT_Worksheet's 38-child sequence:
   `conditionalFormatting*` then `dataValidations` slot BETWEEN `mergeCells` and
   `hyperlinks`; `tableParts` slots AFTER `legacyDrawing`, just before `extLst` — NOT
   after hyperlinks. Writer's order comment (writer/sheet.ts) gets the full sequence;
   each insertion is pinned by a golden-bytes test. (Sheet-level `autoFilter` would slot
   between `sheetData` and `mergeCells` — deferred, see decision 10.)
2. **One shared model per feature** in `types.ts` (reader returns IS writer input; bridge
   = structural pass-through): `TableInfo` (name/ref/header/totals/columns/styleInfo),
   `DataValidation` (type/operator/formulas/sqref[]/prompts/errors — the `showDropDown`
   attribute is INVERTED in the file format; the model exposes the intuitive boolean and
   documents the inversion), `ConditionalFormattingRule` (discriminated union over
   ST_CfType with raw formulas/cfvos/colors). xlsb/ods/csv readers DEGRADE for all three
   (pinned by tests).
3. **dxf is its own `DxfStyle` type, INLINE in the model — indexes never go public.**
   `DxfStyle` shares Color/FontStyle/BorderEdge primitives but NOT FillStyle semantics: a
   solid dxf fill's visible color is **`bgColor` with patternType absent — the exact
   INVERSE of CellStyle's fgColor rule**. Stored raw, never normalized (a "clean-up" here
   silently swaps every CF highlight color). Inline numFmt code strings (matches our
   model). Emission order inside `<dxf>`: font, numFmt, fill, alignment, border,
   protection. **Ownership:** the reader RESOLVES `@dxfId` → an inline `DxfStyle` on the
   rule/table object at parse time; the writer INTERNS the inline styles into one
   `<dxfs>` table and assigns ids at emit (structural interning, the F4.2 styles
   precedent). No public numeric dxf index anywhere — cfRule `@dxfId` and table `*DxfId`
   share one file-level index space, and pass-through indexes would demand a bridge-wide
   remap whose off-by-one is invisible to schema validation.
4. **x14 posture: degrade on read, never emit.** Worksheet-level `extLst` extensions
   (x14:dataValidations `{CCE6A557-…}`, x14:conditionalFormattings `{78C0D931-…-F0AAD7539E65}`)
   are skipped; the cfRule-level `<extLst>` x14:id twin link `{B025F937-…}` is STRIPPED on
   read (a dangling GUID must not round-trip). Named consequences, documented + pinned:
   Excel-authored cross-sheet DV (x14-only) is invisible; Excel 2010+ dataBars lose their
   x14 twin (render 2007-style). Cross-sheet refs in MAIN-part formula1 are accepted both
   ways (openpyxl/ExcelJS-compatible). The corpus lossless-or-typed property gains x14 as
   an EXPECTED named degradation.
5. **Shared bounds, single-sourced** (reader clamps/drops, writer rejects): DV
   promptTitle/errorTitle ≤32, prompt/error ≤255, inline list literal ≤255; CF priority
   int ≥1; cfvo counts (colorScale 2–3, dataBar 2, iconSet = icon count); `@dxfId` must
   index the emitted dxfs; **sqref stays SYMBOLIC** — parsed with the a1 machinery,
   capped in RANGE COUNT, never expanded per-cell (a whole-grid sqref × 10k ranges is the
   M9 repeat-bomb; F4.4/F4.6 posture).
6. **Writer normalizations (determinism, semantics-preserving):** cfRule priorities are
   renumbered densely 1..n on WRITE **by ascending caller priority, document order as the
   tie-break** — NEVER by document order alone (producer priority order routinely disagrees
   with document order, and priority decides which overlapping rule WINS; renumbering by
   position would silently swap precedence — the exact silent mangling this repo bans).
   The model's `priority` field is therefore honored input (writer validates int ≥1 per
   decision 5, sorts, renumbers densely); the reader surfaces producer priorities verbatim
   — posture snapshotted in the corpus test. Booleans as 1/0, FF-prefixed ARGB out (the
   existing shared `HEX_COLOR` already accepts 6- and 8-digit input on both sides — no
   widening needed or wanted), leading `=` stripped from DV/CF formula text, everything
   through escapeText/isXmlSafe (typed reject), single-read TOCTOU throughout.
7. **Dual-encoded rules round-trip verbatim.** containsText/timePeriod cfRules carry BOTH
   declarative attrs AND a generated `<formula>` (relative to the sqref top-left, often
   containing TODAY()) — both sides pass through untouched; regenerating either
   desynchronizes them. Never evaluate, never rewrite (1904-mode serials included).
8. **Table writer rules (the repair-prompt feature):** tableColumn names DERIVE from
   header-row cell values (single source of truth — openpyxl itself can author
   repair-triggering mismatches; reject empty/duplicate/non-string headers typed);
   workbook-unique auto-assigned `@id`; displayName uniqueness CASE-INSENSITIVE across
   sheets + lexical rules (no spaces, not a cell ref incl. bare C/R, ≤255) in a shared
   constants module (a future defined-names feature reads the same bounds); tableColumns
   count == ref width; overlapping table ranges rejected; `autoFilter` child written by
   default; rel + content-type Override only when tables exist (byte-identity).
9. **Byte-identity.** A workbook using none of the M9 features emits EXACT pre-M9 bytes —
   no tables dir, no Overrides, no empty `<dataValidations>`/`<dxfs>`. Corpus property +
   bridge snapshot extend to tables/DV/CF.
10. **Scope fences (named):** no sheet-level autoFilter authoring (needs the hidden
    `_xlnm._FilterDatabase` defined name — defer); no custom table-style DEFINITION parts
    (built-in `tableStyleInfo` names only); iconSet limited to built-in ST_IconSetType;
    write-side x14 never; LibreOffice models tables as database ranges — a divergence to
    test when an LO fixture exists, not to absorb.
11. **Both writers, same features.** `streamXlsx` shares the buffered writer's element
    templates but assembles its own header/footer strings — DV/CF slot into the streamed
    footer exactly as into the buffered template, and tables emit their side parts through
    the shared `sheetSideParts` plumbing (F6.1 precedent). Every F9.1–F9.3 feature lands
    on BOTH writers in the same commit, with the streamed==buffered equivalence test
    extended — a streaming writer that silently DROPS a sheet's validations is the failure
    mode this line exists to prevent.
12. **Oracles:** openpyxl 3.1.5 (verified: reads+writes all three, round-trips
    warnings-as-errors clean; emits NO x14 → exercises exactly the base schema). ExcelJS
    4.4.0 (already a bench devDep) is the ONLY local x14 producer — authors the
    x14-degrade fixtures. SheetJS = loads-without-throwing smoke only. **Real Excel =
    the owner** — the repair-prompt property (element order, header/name match) can only
    be truly verified there; named dependency, not silently dropped.

**Standing gates** as always; every writer-touching feature runs the byte-identity recipe.

**Dependency order:** F9.1 tables → F9.2 data validation → F9.3 conditional formatting +
dxfs (largest; consumes the dxf index rules) → F9.4 fuzzing + corpus + milestone review.

### F9.1 — Tables: read + write + bridge ☑
**Scope (in).** Read: `xl/tables/tableN.xml` (`@ref @displayName @headerRowCount
@totalsRowCount`, `autoFilter`, `tableColumns` (+`@totalsRowFunction/@totalsRowLabel`,
`totalsRowFormula` for custom), `tableStyleInfo`) via the sheet's `tableParts` rels;
`Worksheet.tables: readonly TableInfo[]`. Write: `SheetInput.tables` per decision-8 rules;
part + rel (`…/relationships/table`) + content-type Override; `<tableParts>` after
`legacyDrawing`. Bridge carries. Table dxf attrs (`headerRowDxfId` etc.) are DROPPED
NAMED in F9.1 (no public raw indexes — decision 3); when F9.3's dxf parser lands they
become inline `DxfStyle` fields on `TableInfo` columns, additively.
**Scope (out).** calculatedColumnFormula authoring (read verbatim, carried); totals-row
CELL formulas (Excel writes SUBTOTAL(109,…) cells — we write the table part only,
documented); structured-ref formula rewriting.
**Tasks**
- [x] Table part parser (`ooxml/table.ts` `parseTable`) + `TableInfo`/`TableColumn`/
      `TableStyleInfo` (types.ts) + `Worksheet.tables` accessor + degrade pins (ods/xlsb/csv → `[]`).
- [x] Writer: header-derived column names, workbook-global id/name context, part/rel/CT wiring,
      `<tableParts>` order pin (after `<legacyDrawing>`), byte-identity, BOTH writers (decision 11).
- [x] Bridge + corpus snapshot extension; fixtures `openpyxl-tables.xlsx` + crafted edge cases
      (header mismatch → typed reject; duplicate displayName; overlap — all in table-write.test.ts).
- [x] Oracle both ways (openpyxl reads OUR output warnings-as-errors clean + we read openpyxl- and
      Excel-authored tables). **Owner confirmed the repair-prompt check in real Microsoft Excel
      (decision 12): an openjsxl-written `Inventory` table opens clean — filter dropdowns, banded
      rows (`TableStyleMedium9`), no repair dialog.**
- [x] Adversarial review (4 lenses + refuting verifiers, openpyxl-oracle-wired).
**Acceptance.** inventory-table.xlsx reads verbatim; write→openpyxl clean; no-table
workbooks byte-identical; every decision-8 rejection typed + named.

**Landed (F9.1, committed `c4d191b`).** `ooxml/table.ts` (tolerant `parseTable`)
+ `TableInfo`/`TableColumn`/`TableStyleInfo` (shared model) + `Worksheet.tables` (reader accessor,
lazy; ods/xlsb/csv degrade to `[]`). Writer: `buildTables` (sheet.ts) validates + emits per decision
8 — column names DERIVE from the header row (buffered) or `columns[].name` (streaming, which can't
read the streamed header); workbook-global id/part-number + case-insensitive-unique display names via
`createTableContext`; ref width == column count; overlapping ranges rejected; typed `XlsxError` on
every violation. `<tableParts>` slots LAST in CT_Worksheet (after `<legacyDrawing>`); part + rel +
content-type Override emitted only when a sheet has tables. **Both writers** (`writeXlsx` +
`streamXlsx`) wired; the bridge carries tables; corpus property test extended → both real fixtures
round-trip **deep-equal**. Fixtures: `inventory-table.xlsx` (Excel, read verbatim) +
`openpyxl-tables.xlsx` (new, openpyxl-authored). Oracle both ways: openpyxl reads our output
warnings-as-errors clean; our reader reads openpyxl/Excel tables.
- **Adversarial review — 11 findings, 4 CONFIRMED (0 unverified), all fixed + pinned:** a
  `totalsRow` on a too-short ref drove the auto-filter end row below the start row → a **bare**
  `Error("invalid row index: 0")` (violating the typed-rejection + corpus-property contract) or a
  reversed `<autoFilter>` range (3 findings, one root cause) → now a typed `XlsxError` reject; and
  `<autoFilter>` was emitted for header-LESS tables, which Excel/openpyxl never do → now omitted when
  `headerRow` is false. Also proactively closed a single-read TOCTOU on `columns[c]`.
**Gate:** biome 0 / tsc 0 / **803 tests** (32 new across table.test.ts + table-write.test.ts +
corpus). No-table workbooks byte-identical (golden pins green); the `.xlsx` writer's pre-F9.1 output
is unchanged for input not using tables.

### F9.2 — Data validation: read + write + bridge ☐
**Scope (in).** `<dataValidations>` between mergeCells and hyperlinks. All 8 types +
operators; formula1/formula2 verbatim (incl. cross-sheet in main part); multi-range
`@sqref` (space-separated, symbolic); prompts/errors + `errorStyle`; the dropdown
inversion per decision 2; inline-list quoting quirk (`"a,b,c"` — quotes are formula text)
documented + escaped. x14 DV skipped-named (decision 4). `Worksheet.dataValidations`,
`SheetInput.dataValidations`, bridge carry.
**Tasks**
- [ ] Parser + model + degrade pins; bounds (decision 5) single-sourced.
- [ ] Writer + order pin + byte-identity; escaping/TOCTOU.
- [ ] Fixtures: `openpyxl-datavalidation.xlsx` (all types, multi-range, both dropdown
      states, cross-sheet list) + ExcelJS x14 fixture pinning the degrade.
- [ ] Corpus/bridge extension; oracle both ways; adversarial review (hostile sqref).
**Acceptance.** All-8-types fixture round-trips; x14 degrade pinned; hostile sqref
bounded; DV-free workbooks byte-identical.

### F9.3 — Conditional formatting + differential styles (dxfs) ☐
**Scope (in).** `<conditionalFormatting @sqref>` blocks + `<cfRule>` for the FULL base
ST_CfType set (cellIs, expression, colorScale, dataBar, iconSet, top10, aboveAverage,
uniqueValues/duplicateValues, containsText-family, containsBlanks/Errors-family,
timePeriod) with `@priority/@stopIfTrue/@operator/@text/@rank/@percent/@timePeriod/…`,
`<formula>×0–3`, cfvo/color children (counts per decision 5). `<dxfs>` in styles.xml as
inline `DxfStyle` (decision 3: reader resolves ids at parse, writer interns at emit —
no public index, no bridge remap). Priorities per decision 6; dual-encoding per decision 7; x14 per
decision 4 (incl. rule-level extLst strip; ExcelJS-authored dataBar fixture pins it).
`Worksheet.conditionalFormatting`, `SheetInput.conditionalFormatting`, bridge carry.
**Scope (out).** Rule EVALUATION (which cells currently match is M8-adjacent, out);
custom icon sets; x14 emission.
**Tasks**
- [ ] dxfs parser + writer interning per decision 3 (bgColor rule pinned by a color-swap
      regression; interning correctness pinned — two rules sharing one producer dxf must
      keep identical styles after the round-trip).
- [ ] cfRule parser/emitter (discriminated union; verbatim formulas; the decision-6
      priority posture pinned by a precedence-preservation regression); order pin +
      byte-identity.
- [ ] Retrofit inline `DxfStyle` onto F9.1's `TableInfo` columns (additive).
- [ ] Fixtures: `openpyxl-condformat.xlsx` (≥10 rule types + several dxfs incl.
      font+fill+border+numFmt) + ExcelJS x14 dataBar (degrade pin) + crafted hostile
      (10k ranges × whole-grid sqref, priority collisions/0/negatives).
- [ ] Corpus/bridge extension; oracle both ways; adversarial review.
**Acceptance.** Every base rule type round-trips with its dxf intact (colors NOT swapped);
x14 stripped-named; hostile sqref bounded; CF-free workbooks byte-identical.

### F9.4 — Fuzzing harness, corpus expansion, milestone hardening review ☐
**Context.** jazzer.js (libFuzzer) needs native bindings — out. **fast-check** (MIT, pure
TS, seeded + shrinking, runs inside vitest) is the property-based half; a hand-rolled
seeded mutation engine (~200 lines, xorshift PRNG — zero deps) is the corpus half. No
maintained pure-JS coverage-guided fuzzer exists (verified July 2026).
**Scope (in).** NEW private `packages/fuzz` (bench precedent; fast-check as its only
devDep). Half A — properties over WRITER inputs: generated workbook trees incl. hostile
getters/Proxies, non-plain prototypes, unknown keys; assert `writeXlsx` round-trips
through `openXlsx` OR throws `XlsxError('invalid-input')` only, plus write-twice
byte-identical determinism. Half B — seeded mutation over `packages/fixtures/data/*`:
bit/byte flips, truncation, zip-structure mutations (EOCD/central-directory counts,
offsets, CRCs, name lengths), XML-aware mutations (huge count attrs, sqref explosion,
deep nesting, duplicate attrs, encoding garbage), replayed against all four openers under
wall-clock + RSS budgets; invariant: resolve OR `instanceof XlsxError` — any
TypeError/RangeError/OOM/timeout is a crasher. **CI posture:** a fixed-seed smoke
(fast-check numRuns ≈50 + a few hundred mutants, <10 s) lives as an ordinary vitest suite
in `packages/fuzz/__tests__` so the root gate covers it with zero new plumbing; long runs
are local-only (`pnpm --filter @openjsxl/fuzz fuzz -- --ms=…`). **Triage pipeline:**
crashers → gitignored `packages/fuzz/crashers/` (reproducer bytes + seed + trace),
auto-minimized (chunk-removal binary search), fixed, then MANUALLY promoted to
`packages/fixtures/data/edge-*.…` with provenance + a verbatim-read regression (the
existing data/README checklist).
**Corpus expansion.** Authorable now: the three openpyxl fixtures (F9.1–F9.3) + the
ExcelJS x14 fixture. **Owner-provided asks (named, not silently dropped):** a real Excel
365 file (negative-value dataBar + x14 twin, custom icon set, cross-sheet DV in extLst,
calculatedColumnFormula table); a Google Sheets export with DV+CF; a LibreOffice file
(tables-as-database-ranges divergence). Differently-licensed → gitignored `local/`.
**Tasks**
- [ ] `packages/fuzz` skeleton + fast-check writer properties (Half A).
- [ ] Mutation engine + four-opener replay harness + budgets (Half B).
- [ ] CI smoke wiring (fixed-seed vitest suite; flake-proof: generous budgets, no
      wall-clock asserts in CI — those live in the local long-run only).
- [ ] Triage/minimize/promote pipeline + crashers/ gitignore from day one.
- [ ] Corpus: author + commit the fixtures above; file the owner asks.
- [ ] Long local run against M9-complete readers; fix + pin every crasher; full-milestone
      adversarial review (cross-feature + hardening lenses).
**Acceptance.** The smoke runs in the standard gate deterministically; a long local run
completes with zero unexplained crashers (every one fixed + promoted or explained); the
corpus property covers tables/DV/CF; M9 done = 0.9 release prep, owner-gated.

---

## M10+ — Later milestones (outline; expanded when reached)

- **Deferred — legacy `.xls` (BIFF8) read:** a CFB/OLE2 container reader + BIFF record
  layer (globals + per-sheet substreams; the SST `Continue`-split Unicode-flag trap;
  `FILEPASS`/`VelvetSweatshop` encryption → typed reject). Cut from M7 (owner, 2026-07-10)
  as the milestone's biggest lift for a declining ~1997–2007 install base. Revisit if
  users ask; the multi-format seam (F7.1) already accommodates another `open*` + backend.
- **Deferred — native lane:** the optional `@openjsxl/native` napi-rs binding to Rust
  `calamine` (and a WASM build) behind the zip/xml interface. Deferred by the F5.5
  benchmark evidence (see the M6 re-scope note); revisit if a workload shifts the math.
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
