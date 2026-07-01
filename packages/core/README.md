# @openjsxl/core

The zero-dependency OOXML engine behind [`openjsxl`](https://www.npmjs.com/package/openjsxl):
the `zip → xml → ooxml → reader` layers that turn an `.xlsx` into typed cells, built only on
platform Web APIs (`DecompressionStream`, `TextDecoder`, …). No runtime dependencies.

**Most users should install [`openjsxl`](https://www.npmjs.com/package/openjsxl) instead** — it
re-exports everything here and is the stable public surface. Install `@openjsxl/core` directly
only if you want the engine without the facade.

```sh
npm install @openjsxl/core
```

## Usage

```ts
import { openXlsx, streamSheetRows, XlsxError } from '@openjsxl/core'
import { readFile } from 'node:fs/promises'

const wb = await openXlsx(await readFile('data.xlsx'))
const sheet = wb.sheet('Sheet1')

sheet.cell('A1') // { ref, type, value } — narrow on `type` for a typed value
sheet.numberFormat('C1') // "mm-dd-yy" | undefined
sheet.mergedCells // ["A1:B1", …]

// Constant-memory streaming for large sheets — one row at a time.
for await (const row of streamSheetRows(await readFile('huge.xlsx'))) {
	console.log(row.index, row.cells.length)
}
```

Malformed input throws a typed `XlsxError` with a discriminating `.code`
(`'not-a-zip' | 'not-xlsx' | 'missing-part' | 'corrupt-zip' | 'part-too-large' | …`), never a
bare `TypeError` from a corrupt file.

## Exports

- **Reader:** `openXlsx`, `streamSheetRows`, `Workbook`, `Worksheet`, `ReadOptions`
- **Errors:** `XlsxError`, `XlsxErrorCode`
- **Types:** `Row`, `Cell`, `CellType`, `Comment`, `Hyperlink`, `SheetInfo`, `CellRef`
- **A1 & dates:** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef`, `serialToDate`

Full guide, design notes, and roadmap: <https://github.com/joaquimserafim/openjsxl>

## License

MIT
