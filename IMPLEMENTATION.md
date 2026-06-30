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

### F1.5 — Shared strings table ☐
**Context.** Most text lives once in `xl/sharedStrings.xml`; cells of type `s` store a
zero-based index into it.
**Scope.** `parseSharedStrings(xml)` → `string[]`.
**Design notes.** Parse once. Handle plain `<si><t>` and rich-text `<si><r><t>…</t></r>`
(concatenate runs). Respect `xml:space="preserve"`. This is the first big consumer of the
streaming tokenizer.
**Tasks**
- [ ] `ooxml/shared-strings.ts` (plain + rich-text runs).
- [ ] `ooxml/shared-strings.test.ts`.
**Acceptance.** Index → string matches the fixture's `sst` exactly, runs concatenated.

### F1.6 — Worksheet cell stream & typing ☐
**Context.** The core value: turn `<c>` elements into typed `Cell`s.
**Scope.** A worksheet streamer that yields rows of `Cell`; `decodeCell(raw, ctx)` mapping
the `t` attribute + content to the discriminated union.
**Design notes.** `t` values: absent/`n` = number (default), `s` = shared-string index,
`str` = cached formula string, `b` = boolean (`0`/`1` inside `<v>`), `e` = error,
`inlineStr` = text in `<is><t>`. Booleans/errors live **inside `<v>`** — don't read them as
numbers. Cells/rows are **sparse and may be unordered** — key by ref. Date detection is
deferred to F2.1 (until then a date-styled number reads as a number).
**Tasks**
- [ ] `ooxml/cell.ts` `decodeCell` for all `t` variants (date typing stubbed to number).
- [ ] `reader/worksheet` streamer over `sheetN.xml` using the tokenizer.
- [ ] Tests: each `t` variant from a fixture; sparse rows; out-of-order cells.
**Acceptance.** Reads string/number/bool/error/inline/formula-cached values with correct
types from `basic.xlsx`.

### F1.7 — Reader public API ☐
**Context.** The ergonomic surface users actually touch.
**Scope.** `openXlsx(source)` → `Workbook`; `Workbook.sheets`, `Workbook.sheet(name)`;
`Worksheet.cell('A1')`, `Worksheet.rows()` async iterator. `source`: `Uint8Array | ArrayBuffer`.
**Design notes.** Async (decompression is async). Lazy per-sheet parsing. TS
discriminated-union cells so `cell.type === 'date'` narrows `cell.value`.
**Tasks**
- [ ] `reader/workbook.ts` wiring zip → rels → sharedStrings → worksheet.
- [ ] `cell('A1')` and `for await (const row of sheet.rows())`.
- [ ] `reader/index.ts` + export from `core/src/index.ts` + facade.
**Acceptance.** `(await openXlsx(bytes)).sheet('Sheet1').cell('A1').value` is the right
typed value.

### F1.8 — Vertical-slice tests ☐
**Context.** Lock the MVP behavior with end-to-end tests on a real file.
**Tasks**
- [ ] Wire `@openjsxl/fixtures` into the reader tests (alias + tsconfig path).
- [ ] Un-`todo` the slice: list sheets via rels; read string/number/bool; (date after F2.1).
- [ ] README quick-start: "`.xlsx` → JSON in < 50 LOC".
**Acceptance.** `pnpm test` exercises a real `.xlsx` end-to-end and is green. **Tag v0.1.**

---

## M2 — Reader hardening (v0.2)

Goal: correct dates, constant memory, and the common metadata real files carry.

### F2.1 — Styles & number-format date detection ☐
**Context.** Nothing on a cell says "date" — a number is a date iff its style applies a
date/time number format. This unlocks correct `date` cells.
**Scope.** `parseStyles(xml)` → `StyleTable.isDateStyle(styleIndex)`; integrate into
`decodeCell` so date-styled numbers become `Date` (via F0.4).
**Design notes.** `c@s` indexes `cellXfs`; each `xf` → `numFmtId`. IDs < 164 are **built-ins
not listed in `<numFmts>`** (hardcode the date/time ones: 14–22, 45–47, etc.); IDs ≥ 164 are
custom in `<numFmts>` — detect date tokens (`y m d h s`) outside quoted literals.
`cellXfs[0]` is the implicit default (omitted `s` ⇒ `s=0`).
**Tasks**
- [ ] `ooxml/styles.ts` (cellXfs + numFmts + built-in date table + format-code date sniff).
- [ ] Integrate into `decodeCell`; honor `date1904`.
- [ ] Tests across built-in and custom date formats + a non-date number with `s`.
**Acceptance.** `C1` of `basic.xlsx` reads as a `Date`; a plain number with a style stays a
number.

### F2.2 — Constant-memory streaming reader ☐
**Context.** Worksheets and shared strings can be huge; peak memory must not scale with file
size.
**Scope.** Stream `sheetN.xml` and `sharedStrings.xml` from the zip without materializing
whole parts; expose row-at-a-time iteration end-to-end.
**Design notes.** Drive `DecompressionStream`'s readable directly into the tokenizer;
chunk-boundary-safe scanning. Optional interned shared-strings to cap memory.
**Tasks**
- [ ] Streaming inflate → tokenizer bridge (handle split tokens across chunks).
- [ ] Streaming shared-strings (or interned lookup).
- [ ] Memory test: read a large generated sheet under a fixed cap.
**Acceptance.** Reads a 100k-row generated sheet with roughly constant memory.

### F2.3 — Common metadata ☐
**Scope.** Merged cells, hyperlinks, comments, per-cell number-format strings, sheet
visibility/dimensions.
**Tasks**
- [ ] Merged ranges; hyperlinks (via worksheet rels); comments; `numberFormat` on cells;
      `SheetInfo.visible`/dimensions. Each fixture-backed.
**Acceptance.** Each feature verified against a real-producer fixture.

### F2.4 — Robustness ☐
**Scope.** Sparse/unordered cells, missing `<dimension>`, large/odd shared strings,
malformed input.
**Tasks**
- [ ] Tolerate missing optional parts; clear typed errors on corrupt input; fuzz the
      tokenizer/zip parser.
- [ ] Zip hardening deferred from F1.3: commit a real Excel/LibreOffice/Sheets `.xlsx`
      fixture (closes the F1.3 review's S7); a configurable max part size; decide policy for
      duplicate entry names (last-wins today) and directory entries (`name/` placeholders).
**Acceptance.** No crashes on a corpus of malformed/edge-case files; errors are actionable.

---

## M3 — Writer MVP (v0.3)

Goal: write a valid `.xlsx` Excel opens without "repair", round-tripping values and types.

### F3.1 — ZIP / OPC writer ☐
**Scope.** Build a ZIP from parts: stored and/or `CompressionStream('deflate-raw')`, CRC32,
correct local/central headers + EOCD.
**Design notes.** Reuse the generator's stored-zip code (F0.2) as the basis; add deflate for
large parts. Prefer **inline strings** when writing to avoid buffering a shared table.
**Tasks**
- [ ] `writer/zip.ts` (store + deflate + CRC32); round-trip with `openZip`.
**Acceptance.** Output re-reads byte-identically through F1.3.

### F3.2 — Minimal workbook writer ☐
**Scope.** Emit the minimal valid part set: `[Content_Types].xml`, `_rels/.rels`,
`xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/worksheets/sheet1.xml`. Values, types
(string/number/bool/date), multiple sheets.
**Design notes.** Escape `& < >`; `xml:space="preserve"` for edge whitespace; clamp floats to
15 significant figures; dates as serials + a date style; no `calcChain.xml`.
**Tasks**
- [ ] `writer/workbook.ts` + public `openjsxl/write` entry; create sheets, set cells, save.
**Acceptance.** A written file opens cleanly in Excel **and** LibreOffice.

### F3.3 — Round-trip fidelity ☐
**Tasks**
- [ ] Write → read → assert values/types/sheets; golden-file diffs.
**Acceptance.** Round-trip is lossless for the supported value set. **Tag v0.3.**

---

## M4+ — Later milestones (outline; expanded when reached)

- **M4 — Styles (v0.4):** read + write fonts, fills, borders, alignment, named/number
  formats, column widths / row heights / freeze panes. *The SheetJS-Pro killer.*
- **M5 — Streaming writer + native lane (v0.5):** constant-memory writer; images; an
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
