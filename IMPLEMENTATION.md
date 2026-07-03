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

### F5.5 — Benchmark harness + published numbers ☐
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
- [ ] harness + workload generator (reuses `@openjsxl/fixtures` builder); runner with
      warmup/median; markdown reporter.
- [ ] first published `docs/benchmarks.md` + README link with date + hardware note.
**Acceptance.** `pnpm bench` reproduces the table end-to-end on a clean checkout; README
claims match the published numbers (no unmeasured "fast" claims anywhere in docs).

---

## M6+ — Later milestones (outline; expanded when reached)

- **M6 — Images + native lane (v0.6):** picture read (drawingML anchors → `{ref, bytes,
  mime}`) and basic image write; the **optional** `@openjsxl/native` napi-rs binding to
  the Rust `calamine` reader (and a WASM build) behind the zip/xml interface, selected via
  `optionalDependencies` with the pure-TS path as universal fallback — justified (or not)
  by the F5.5 numbers.
- **M7 — More formats (v0.7):** `.xlsb` (binary) read; `.ods` read; legacy `.xls`
  (BIFF8) read.
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
