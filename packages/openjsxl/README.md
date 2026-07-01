# openjsxl

Fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) reader for JavaScript runtimes —
Node, Deno, Bun, the browser, and edge. This is the package to install; it re-exports the
[`@openjsxl/core`](https://www.npmjs.com/package/@openjsxl/core) engine.

```sh
npm install openjsxl
```

```ts
import { openXlsx } from 'openjsxl'
import { readFile } from 'node:fs/promises'

const wb = await openXlsx(await readFile('data.xlsx'))

// Typed cells: narrowing on `cell.type` gives a correctly typed `cell.value`.
const a1 = wb.sheet('Sheet1').cell('A1')
console.log(a1.type, a1.value) // e.g. "string" "hello"

// Stream a whole sheet, row at a time.
for await (const row of wb.sheet(wb.sheets[0].name).rows()) {
	console.log(row.index, row.cells.map((c) => c.value))
}
```

For very large sheets use `streamSheetRows` (constant memory); a worksheet also exposes
`numberFormat`, `dimension`, `mergedCells`, `hyperlinks`, `comments`, and `visible`, and the
reader throws a typed `XlsxError` (with a discriminating `code`) on malformed input.

See the [project README](https://github.com/joaquimserafim/openjsxl#readme) for the full guide,
design notes, and roadmap.

## License

MIT
