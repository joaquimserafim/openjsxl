# @openjsxl/core

The zero-dependency OOXML engine behind [`openjsxl`](https://www.npmjs.com/package/openjsxl) —
the `zip → xml → ooxml → reader` layers that turn an `.xlsx` into typed cells, built only on
platform Web APIs (`DecompressionStream`, `TextDecoder`, …).

**Most users should install [`openjsxl`](https://www.npmjs.com/package/openjsxl) instead** — it
re-exports everything here and is the stable public surface. Install `@openjsxl/core` directly
only if you want the engine without the facade.

```sh
npm install @openjsxl/core
```

```ts
import { openXlsx, streamSheetRows, XlsxError } from '@openjsxl/core'
```

The API (`openXlsx`, `streamSheetRows`, `Workbook`, `Worksheet`, `XlsxError`, and the A1/date
helpers) is documented in the [project README](https://github.com/joaquimserafim/openjsxl#readme).

## License

MIT
