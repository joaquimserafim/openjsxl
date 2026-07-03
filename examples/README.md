# Examples

Runnable usage examples for [`openjsxl`](https://www.npmjs.com/package/openjsxl). Each script
imports from `openjsxl` exactly as an installed consumer would, reads
[`data/sample.xlsx`](./data/sample.xlsx), and prints what the library sees.

`sample.xlsx` is a small two-sheet workbook: a visible **Sales** sheet (typed cells, a
date-formatted column, a cached formula, a merged range, a hyperlink, and a comment) and a
hidden **Archive** sheet.

## Run

From the repo root, once:

```sh
pnpm install   # links the workspace `openjsxl` into ./examples/node_modules
```

Then any example:

```sh
node examples/01-read-cells.mjs      # typed cell access
node examples/02-sheet-to-json.mjs   # sheet → JSON records
node examples/03-stream-rows.mjs     # constant-memory streaming
node examples/04-metadata.mjs        # number formats, merges, links, comments, visibility
node examples/05-error-handling.mjs  # typed XlsxError on bad input
node examples/06-write.mjs           # write .xlsx + read → modify → write
node examples/07-styles-and-layout.mjs # styled cells, widths, freeze, merges, links (0.4)
node examples/08-streaming-write.mjs # constant-memory write from an async row source (0.5)
node examples/09-comments-formulas-theme.mjs # comments, live formulas, resolveColor (0.5)
```

Or all of them: `pnpm --filter openjsxl-examples all`.
