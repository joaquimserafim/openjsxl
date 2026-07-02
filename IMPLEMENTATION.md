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

### F4.3 — Number-format write (built-in reverse map + custom ids ≥ 164) ☐
**Scope.** Activate `CellStyle.numberFormat` in the writer as a format **code string**. Codes
exactly matching `BUILTIN_FORMATS` reverse-map to their id (no `<numFmts>` entry); others intern
from 164 up in deterministic first-encounter order. A `Date` with a user code keeps it (implicit
id 14 only when absent). Re-read typing flows through `isDateFormatCode` — a number written with
a date code re-reads as `date` (Excel-faithful; documented).
**Tasks**
- [ ] `BUILTIN_CODE_TO_ID` reverse map; registry numFmt interning; `<numFmts>` emission.
- [ ] Tests: built-in maps flat; custom round-trips verbatim via `numberFormat(ref)`; date
  interplay both directions; locale-id codes documented as non-representable.
**Acceptance.** `numberFormat(ref)` returns the written code verbatim after round-trip.

### F4.4 — Bridge carries styles (round-trip fidelity) ☐
**Scope.** `workbookToInput` attaches `sheet.style(cell.ref)` per populated cell — `{value,
style}` only when a style exists (unstyled v0.3-era workbooks produce identical input → the
byte-identical path), including styled *empty* cells (`<c s/>`) it currently drops. Signature
unchanged. README fidelity table rows move to "lossless".
**Design notes.** Documented flattening: row/column-default styles resolve into per-cell styles;
files authored under a **custom theme** keep `{theme,tint}` indices but re-render against our
default theme after rewrite — documented loudly. Property test: bridge output must always pass
`writeXlsx` validation.
**Tasks**
- [ ] Bridge style attachment (incl. styled empties); golden pins; Excel/openpyxl-authored
  styled fixtures round-trip; README fidelity table update.
**Acceptance.** read → bridge → write → read gives deep-equal `style(ref)` for the supported
style set; unstyled files byte-identical.

### F4.5 — Sheet geometry: column widths, row heights, hidden, freeze panes ☐
**Scope.** Read: three lazy accessors in the `mergedCells` idiom — `columns`
(`{min,max,width?,hidden?}` from `<cols>`), `rowProperties` (`Map<row, {height?,hidden?}>` from
`<row>` attrs, dedicated scan off the hot path), `freeze` (`<pane state="frozen">`; split panes
read `undefined`). Write: matching `SheetInput.columns/rowProperties/freeze`; schema order
`sheetViews` → `cols` → `sheetData`; property-only rows emit cell-less `<row>`.
**Tasks**
- [ ] Reader accessors + types; writer emission + validation; bridge carries geometry.
- [ ] Tests: real-producer fixture reads exact widths/heights/hidden/freeze; write → re-read
  equal; openpyxl confirms frozen pane + widths.
**Acceptance.** Geometry round-trips; Excel/LibreOffice show frozen header and sized columns.

### F4.6 — Structural metadata write: merges, hyperlinks, visibility ☐
**Scope.** `SheetInput` gains `merges` (A1 ranges → `<mergeCells>`; malformed/single-cell/
overlapping rejected — Excel repair-prompts on overlap), `hyperlinks` (reader-mirroring records →
`<hyperlinks>` + the writer's **first per-sheet rels part**, `TargetMode="External"` for
targets), and `state: 'visible'|'hidden'|'veryHidden'` (≥1 sheet must stay visible). Reader
`SheetInfo` gains `state` additively (`visible` boolean retained). Bridge carries all three.
**Tasks**
- [ ] Merges + validation; hyperlinks + per-sheet rels wiring; visibility + guard;
  `SheetInfo.state`; bridge; tests (re-read via `mergedCells`/`hyperlinks`/`state`; openpyxl +
  Excel agree, links resolve, no repair prompt).
**Acceptance.** The bridge's v0.3 drop-list shrinks to: comments, formulas, error cells.
**Release (separate tail commit): tag v0.4** — bump `0.4.0`, README fidelity table + styled
example, PUBLISHING note.

---

## M5+ — Later milestones (outline; expanded when reached)

- **M5 — Streaming writer + native lane (v0.5):** constant-memory writer; comments write (VML
  legacy drawing); theme1.xml parse + resolved-rgb color helper; images; an
  **optional** `@openjsxl/native` napi-rs binding to the Rust `calamine` reader (and a WASM
  build) behind the F-layer interface, selected via `optionalDependencies` with pure-TS
  fallback.
- **M6–M7 — More formats (v0.6–0.7):** `.xlsb` (binary) read; `.ods` read; legacy `.xls`
  (BIFF8) read.
- **M8 — Formulas (v0.8):** opt-in formula parser + evaluator for common functions, behind
  a separate entry point so the core stays lean.
- **M9 — Breadth + hardening (v0.9):** tables, data validation, conditional formatting;
  fuzzing, expanded corpus, perf benchmarks.
- **1.0:** frozen API, full round-trip fidelity, documentation site, published benchmarks vs
  SheetJS/ExcelJS/openpyxl/calamine.

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
