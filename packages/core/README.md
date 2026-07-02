# @openjsxl/core

[![npm version](https://img.shields.io/npm/v/@openjsxl/core?color=cb3837&logo=npm)](https://www.npmjs.com/package/@openjsxl/core)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen)](https://www.npmjs.com/package/@openjsxl/core?activeTab=dependencies)
[![license: MIT](https://img.shields.io/npm/l/@openjsxl/core?color=blue)](./LICENSE)

The zero-dependency OOXML engine behind [`openjsxl`](https://www.npmjs.com/package/openjsxl):
the `zip → xml → ooxml → reader` layers that turn an `.xlsx` into typed cells — and the mirror
`writer` layer that turns plain data back into `.xlsx` bytes. Built only on platform Web APIs
(`DecompressionStream`, `CompressionStream`, `TextDecoder`, …). No runtime dependencies.

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

## Writing

```ts
import { writeXlsx, workbookToInput } from '@openjsxl/core'

// Author from plain data — cell types inferred from the JS values.
const bytes = await writeXlsx({
	sheets: [{ name: 'Report', rows: [['Item', 'Added'], ['Apples', new Date('2024-01-15')]] }],
})

// Or read → modify → write.
const input = await workbookToInput(await openXlsx(bytes))
input.sheets[0].rows.push(['Pears', new Date('2024-02-01')])
const updated = await writeXlsx(input)
```

Output is deterministic; unrepresentable input (no sheets, bad/duplicate sheet name, non-finite
number, invalid `Date`, XML-illegal characters) throws `XlsxError` with `code: 'invalid-input'`.
The round trip is lossless for values, types, and sheet names/order.

## Exports

- **Reader:** `openXlsx`, `streamSheetRows`, `Workbook`, `Worksheet`, `ReadOptions`
- **Writer:** `writeXlsx`, `workbookToInput`, `WorkbookInput`, `SheetInput`, `CellValue`, `WriteOptions`
- **Errors:** `XlsxError`, `XlsxErrorCode`
- **Types:** `Row`, `Cell`, `CellType`, `Comment`, `Hyperlink`, `SheetInfo`, `CellRef`
- **A1 & dates:** `columnToIndex`, `indexToColumn`, `parseRef`, `formatRef`, `serialToDate`

Full guide, design notes, and roadmap: <https://github.com/joaquimserafim/openjsxl>

## License

MIT
