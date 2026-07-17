# How openjsxl is built

A short tour of the internals and the decisions behind them. For usage, see the
[README](./README.md).

## The shape

openjsxl is a small stack of layers, each depending only on the one below it:

```
zip    deflate / inflate on the platform's Compression/DecompressionStream — no zlib binding
xml    a streaming SAX tokenizer, never a DOM
ooxml  the spreadsheet model: workbook, sheets, styles, shared strings, A1 refs, dates, bounds
reader / writer   the public API
```

The reader walks an `.xlsx` from the ZIP central directory inward, inflating each part on
demand and tokenizing its XML in chunks. The writer is the mirror image — it turns plain data
into the same parts and deflates them back into a ZIP. Every byte the writer emits reads back
through the reader, and that round trip is the core of the test suite.

## Principles

- **Zero runtime dependencies.** Everything comes from platform Web APIs —
  `DecompressionStream` / `CompressionStream` for ZIP, `TextEncoder` / `TextDecoder` for
  strings. Nothing is pulled from npm at runtime, which is what keeps the install tiny and the
  supply chain empty.

- **Deterministic bytes.** The same input always produces the same output — no timestamps, no
  ordering wobble, no randomness. Two writes of one workbook are byte-identical, so output is
  diffable and easy to cache.

- **The reader is tolerant; the writer is strict.** Real files from other tools can be messy,
  so the reader clamps or drops out-of-range values and quietly degrades features it can't
  represent instead of throwing. The writer refuses to emit anything a spreadsheet app would
  reject, with a typed error. Both sides share the *same* bounds — whatever the reader hands
  you, the writer accepts, or it fails loudly.

- **One model, both ways.** What a reader accessor returns is exactly what the writer takes.
  There is no separate "write" flavor of a style or a merge, so read → tweak → write is a
  straight pass-through and anything you don't touch round-trips unchanged.

- **Streaming where it counts.** A huge sheet never has to live in memory: rows can be read one
  at a time and written from a lazy (even async) source, holding memory roughly flat no matter
  the row count.

- **Layered and swappable.** The hot path sits behind a narrow interface, so a native
  (Rust / WASM) backend could slot in later without disturbing the layers above it.

## Formats

openjsxl **writes** `.xlsx`, and **reads** more: `.xlsx` / `.xlsm`, `.xlsb`, `.ods`, and
`.csv` / `.tsv`, all into the same typed `Workbook`. A format that can't express a feature
degrades on that accessor rather than throwing — so "someone uploaded a spreadsheet" is one
code path, and any reader becomes a converter to `.xlsx` through the bridge.

## Repo layout

```
packages/core/src/
  zip/      ZIP reader (central-directory walk + inflate) · writer (deflate + headers)
  xml/      SAX tokenizer, plus a chunk-safe streaming variant
  ooxml/    workbook · shared strings · styles · A1 refs · dates · bounds — the OOXML model
  ods/      OpenDocument reader     biff/ + xlsb/  binary .xlsb layers     csv/  delimited text
  reader/   openXlsx / openXlsb / openOds / openCsv / detect + the Workbook / Worksheet API
  writer/   writeXlsx + the bridge that turns a read workbook back into writer input
  formula/  the opt-in formula parser + evaluator (a separate entry point)
packages/openjsxl/   the facade you install (re-exports core)
examples/            runnable, copy-pasteable usage
```

The public API grows only through `packages/core/src/index.ts`; the facade re-exports it.
