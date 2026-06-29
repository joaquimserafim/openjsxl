# openjsxl

A fast, **zero-dependency**, TypeScript-first Excel (`.xlsx`) library for JavaScript
runtimes — Node, Deno, Bun, the browser, and edge.

> Status: **pre-alpha.** The repo is being built in the open, plan-first. There is no
> published package yet. See the [roadmap](./ROADMAP.md) and [implementation
> plan](./IMPLEMENTATION.md).

## Why

JavaScript has no Excel library that is, all at once, maintained, permissively licensed,
published on npm, ESM/TypeScript-first, dependency-free, and fast. SheetJS is frozen on a
vulnerable npm build behind a CDN and a paid tier; ExcelJS is effectively unmaintained.
openjsxl aims for that empty square — taking the speed lessons of Python's
[`python-calamine`](https://pypi.org/project/python-calamine/) and growing toward the
capability of [`openpyxl`](https://pypi.org/project/openpyxl/).

## Approach

- **Read first, write later.** A fast, correct reader earns trust before we ship a writer.
- **Zero runtime dependencies.** Zip inflate/deflate comes from the platform
  (`DecompressionStream` / `CompressionStream`), strings from `TextEncoder`/`TextDecoder`.
- **Layered & swappable.** `zip → xml → ooxml → reader`, with the hot path behind an
  interface so a native (napi-rs / WASM `calamine`) backend can slot in later.

## Packages

| Package | Description |
| --- | --- |
| [`openjsxl`](./packages/openjsxl) | Public facade — the package users install |
| [`@openjsxl/core`](./packages/core) | The zero-dependency OOXML engine |
| `@openjsxl/fixtures` | Private test corpus + generator |

## Development

Requires Node ≥ 24 and pnpm.

```sh
pnpm install
pnpm check       # Biome lint + format check
pnpm typecheck   # tsc, no emit
pnpm test        # Vitest
pnpm fixtures    # regenerate programmatic test fixtures
```

## License

MIT
