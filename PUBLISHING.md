# Publishing

How to cut a release of openjsxl to npm. The repo is a pnpm workspace; releases go out with
`pnpm -r publish`, which handles dependency order and the `workspace:*` rewrite for us.

## What ships

| Package | npm | Notes |
| --- | --- | --- |
| `openjsxl` | ✅ public | The facade users install. Re-exports `@openjsxl/core`. |
| `@openjsxl/core` | ✅ public | The engine. Required — the facade's types re-export it. |
| `@openjsxl/fixtures` | ❌ never | `private: true`. Test corpus; also holds git-ignored Apache-2.0 files under `local/`. |

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

### 1. Preflight — clean build and green gate

```sh
pnpm install
pnpm -r build
pnpm typecheck
pnpm test
```

### 2. Set the version

For the current release the version is `0.9.0` (already set) — skip this step. For a later
release, set the **same** version in both public packages by editing the `"version"` field in:

- `packages/core/package.json`
- `packages/openjsxl/package.json`

(Leave `@openjsxl/fixtures` as-is; it never publishes.) Keep the two in lock-step — the facade
depends on `@openjsxl/core` via `workspace:*`, which is rewritten to that exact version on
publish, so a mismatch would ship a broken dependency range.

### 3. Dry-run — inspect the tarballs

```sh
pnpm -C packages/core     pack
pnpm -C packages/openjsxl pack

tar -tzf packages/core/openjsxl-core-*.tgz
tar -tzf packages/openjsxl/openjsxl-*.tgz

# Facade's core dep must be rewritten (prints "@openjsxl/core": "<version>", not workspace:*):
tar -xzOf packages/openjsxl/openjsxl-*.tgz package/package.json | grep -A2 '"dependencies"'

rm -f packages/*/*.tgz   # clean up
```

Each tarball should contain **only** `package.json`, `README.md`, `LICENSE`, and `dist/*` — no
`src/`, `node_modules/`, `tsconfig`, `.tgz`, or `local/`.

### 4. Publish

Commit everything first — `pnpm publish` refuses an unclean tree (by design).

```sh
git status                        # must be clean
pnpm -r publish --access public   # add --otp=<code> if 2FA is on
```

`pnpm -r publish` publishes `@openjsxl/core` first, then `openjsxl` (topological order), rewrites
`workspace:*` → the version, and skips the private `@openjsxl/fixtures`.

### 5. Verify and tag

```sh
npm view openjsxl version
npm view @openjsxl/core version

git tag -f v<version>             # -f only needed to re-point a tag never published under before
git push -f origin v<version>
```

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
- **`1.0.0`** — bump once the API is settled. Follow semver.

## Notes

- **Private repo links.** While the GitHub repo is private, the `repository`/`homepage`/`bugs`
  URLs in npm metadata will 404 for others. They resolve automatically once the repo is public.
- **Provenance.** To publish with npm provenance from CI later, run `npm publish --provenance`
  from a trusted GitHub Actions workflow; not required for a manual release.
- **Deprecating a bad release.** npm forbids re-publishing a version. If a release is broken,
  publish a patch and `npm deprecate openjsxl@<version> "use <newer>"`.
