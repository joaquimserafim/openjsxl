# openjsxl — Roadmap

High-level direction for **openjsxl**, a fast, zero-dependency, TypeScript-first Excel
(`.xlsx`) library for JavaScript runtimes. This document is the *what* and *why*; the
feature-by-feature breakdown with scoped, trackable tasks lives in
[IMPLEMENTATION.md](./IMPLEMENTATION.md).

---

## 1. Vision

A maintained, MIT-licensed, npm-registry-native Excel library that is **fast and correct
on read** and grows into a **styled writer** — with a modern ESM/TypeScript API that runs
on Node, Deno, Bun, the browser, and edge runtimes.

The name is a deliberate nod to Python's `openpyxl`. The performance bar is set by
`python-calamine` (the Rust reader that pandas and Polars adopted).

## 2. Why now — the gap

| Library | Read / Write | Problem we exploit |
| --- | --- | --- |
| SheetJS `xlsx` | R + W | Frozen on a vulnerable npm build; real releases are CDN-only; styles/charts behind a paid tier |
| ExcelJS | R + W | ~9.7M downloads/week but effectively unmaintained (no release in 2+ years, ~800 open issues) |
| openpyxl (Py) | R + W | The feature benchmark, but pure-Python: slow, ~50× file-size memory |
| python-calamine (Py) | R only | The speed benchmark, but read-only, no styles |

No JavaScript library today is **all** of: maintained, permissively licensed, published on
npm, ESM/TS-first, zero-dependency, and fast. That is the square openjsxl owns.

## 3. Strategy

**Read first (the calamine model), then write (the openpyxl model).**

Reading is where we can be *measurably better*, it is lower-risk (parse a subset and
degrade gracefully), and it is the larger real-world job (ingest, ETL, "user uploaded a
spreadsheet"). A writer that emits a file Excel refuses to open destroys trust instantly,
so it follows once the reader has earned credibility.

Architecturally we copy calamine's core lesson: **be a value extractor, not an object
model** — stream the XML, intern shared strings, return flat typed cells, skip everything
not asked for.

## 4. Principles

1. **Zero runtime dependencies.** Platform built-ins are not dependencies; dev tools
   (TypeScript, Vitest, Biome, tsup) are exempt. This is the anti-SheetJS moat: nothing
   to audit, nothing to CVE, perfect tree-shaking.
2. **Lean on platform primitives.** `DecompressionStream` / `CompressionStream`
   (`deflate-raw`) give us zip inflate/deflate for free; `TextEncoder` / `TextDecoder`
   give us string codecs. Available on Node ≥ 18, Deno, Bun, browsers, and Cloudflare
   Workers.
3. **Layered, bottom-up.** `zip` → `xml` → `ooxml` → `reader`, each independently
   testable, with the hot path (`zip` + `xml`) isolated behind an interface so a native /
   WASM backend can slot in later without touching the rest.
4. **Async-first.** `DecompressionStream` is async; so is `openXlsx()`. This is also the
   right shape for streaming and edge runtimes.
5. **TypeScript-first DX.** Strict types, discriminated-union cell model, ESM only.
6. **Runtime-agnostic.** No Node-only APIs in the core; Node-specific conveniences (path
   I/O, sync inflate) live behind optional entry points.

## 5. Milestones

Status: ☐ not started · ◐ in progress · ☑ done

**Shipped:** everything through **0.6 (Images)** is done and live on npm (`openjsxl` +
`@openjsxl/core` at `0.6.0`, published 2026-07-04). **Next up: 0.7 — more formats.**

| Version | Theme | Outcome | Status |
| --- | --- | --- | --- |
| **M0** | Foundations | Compiling skeleton, fixtures, pure primitives (A1, dates) | ☑ |
| **0.1** | Reader MVP | Read correctly-typed cells (string/number/date/bool) from a real `.xlsx` | ☑ |
| **0.2** | Reader hardening | Constant-memory streaming; styles→date detection; merged cells, hyperlinks, comments | ☑ |
| **0.3** | Writer MVP | Write values, types, sheets; round-trip fidelity tests | ☑ |
| **0.4** | Styles | Read + write fonts, fills, borders, alignment, number formats | ☑ |
| **0.5** | Fidelity + streaming writer | Comments write, formula text, theme fidelity; constant-memory writer; benchmark harness | ☑ |
| **0.6** | Images | drawingML picture read + anchored write; native lane deferred (0.5 benchmarks: pure-TS within ~1.5× of native calamine) | ☑ |
| **0.7** | More formats | `.xlsb`, `.ods`, legacy `.xls` (BIFF8) read | ☐ |
| **0.8** | Formulas | Opt-in formula parse + evaluate (separate entry point; text fidelity ships in 0.5) | ☐ |
| **0.9** | Breadth + hardening | Tables, data validation, conditional formatting; fuzzing + corpus | ☐ |
| **1.0** | Stable | Frozen API, full `.xlsx` round-trip fidelity, docs site, published benchmarks | ☐ |

## 6. How we track

- [IMPLEMENTATION.md](./IMPLEMENTATION.md) breaks every milestone into **features** with
  **context, scope, design notes, tasks (checkboxes), and acceptance criteria**.
- Each feature has a stable ID (e.g. `F1.3`) referenced by branches, commits, and PRs.
- A milestone is "done" when every feature in it is checked and its acceptance tests pass.
