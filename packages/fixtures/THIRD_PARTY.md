# Third-party fixtures

Some real-producer `.xlsx` files under [`data/`](./data) were not exported by us — they are
vendored from other open-source projects so the reader is tested against genuine Excel,
LibreOffice, and openpyxl output (not just our own generator). Each is redistributed under a
license that permits it, recorded below.

## Vendored fixtures

From [**calamine**](https://github.com/tafia/calamine) (`tests/`), MIT-licensed —
Copyright (c) 2016 Johann Tuffe. Each file is an unmodified copy of
`https://github.com/tafia/calamine/blob/master/tests/<name>`:

| File | Producer | Exercises |
| --- | --- | --- |
| `merge_cells.xlsx` | Microsoft Excel | merged ranges (`A1:B1`, `A2:A4`, `B2:D4`) |
| `merged_range.xlsx` | Microsoft Excel | merged ranges, two sheets, per-sheet rels |
| `hyperlinks.xlsx` | openpyxl | external/`mailto:`/internal links, tooltip, display, worksheet rels, `xmlns:r` prefix |
| `date.xlsx` | LibreOffice | custom `numFmt` (`yyyy-mm-dd`, `[hh]:mm:ss`), 1900 date system |
| `date_1904.xlsx` | LibreOffice | `date1904` workbook flag |
| `errors.xlsx` | Microsoft Excel | error cells (`#DIV/0!`, `#NAME?`, `#VALUE!`, `#NULL!`) |
| `inventory-table.xlsx` | Microsoft Excel | shared strings + table (real-world smoke test) |
| `any_sheets.xlsx` | Microsoft Excel | sheet visibility (`hidden`, `veryHidden`) |

The full MIT license text is in [`LICENSE-MIT-calamine.md`](./LICENSE-MIT-calamine.md).

**Reproducing / contributing.** These files are hash-pinned in
[`scripts/fetch-real.mjs`](./scripts/fetch-real.mjs); `pnpm fixtures:real` re-downloads and
verifies them (any upstream change fails loudly). To contribute a real fixture that exercises a
feature end-to-end, add it under `data/`, append an entry (file, url, license, sha256) to the
manifest, and record it in this file — a great first contribution.

## Not committed

Apache-2.0 comment fixtures from [Apache POI](https://github.com/apache/poi)
(`SimpleWithComments.xlsx`, `comments.xlsx`) are useful for local verification but are **not**
committed — they would drag a second license into an otherwise MIT corpus, and as a reader we
only extract `{ref, author, text}`, which a synthetic fixture covers. They live under the
git-ignored `local/` directory; drop them there (or any other producer's file) to test
against them locally. See [`local/README.md`](./local/README.md).
