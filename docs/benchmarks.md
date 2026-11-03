# openjsxl benchmarks

_Measured by `pnpm bench` on 2026-11-03. Re-rendering preserves this measurement date._

**Machine:** Apple M2 Pro · 12 cores · 32 GiB · darwin arm64 · Node v24.22.1

Each number is the **median wall-time** (over warmed iterations) and the **peak process RSS of a
single cold call** (library load + one operation), measured in a **fresh, isolated Node process per
cell** so no library colours another's timing or memory. Workloads are 10 columns wide;
`10k`/`100k`/`1M` count cells. Read fixtures are authored once by ExcelJS (realistic shared
strings + a styles table) and read by all three JS readers. Full methodology and the fairness
caveats are at the bottom.

Libraries: **openjsxl** `1.1.0` (`workspace`); **ExcelJS** `4.4.0` (`4.4.0`; override uuid 11.1.1); **SheetJS** `0.20.3` ([official distribution](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz)). Peak RSS is absolute and includes loading the library.

**Source snapshot:** lockfile SHA-256 `d55cf0f92fcff742726deb66f80e0587234264b1d8dbdf4867b84139a5c59457`; core bundle SHA-256
`d582bc61a4bb4687801062b37c33aecc26705b33da2dc90bf37fdfd7e1a34e22`. The raw result cache also records locked comparator artifact
integrities and fixture hashes.

## Library size

A clean production install (`npm install --omit=dev --ignore-scripts`) of each exact artifact — measured 2026-11-03 on
Node 24.22.1. **Deps** counts the runtime packages the library pulls in
(transitive, excluding the library itself); **Own size** is the package's own unpacked tarball;
**Installed** is the library plus every runtime dependency on disk.
Comparator tarball bytes and npm metadata must match the artifact integrity in the recorded lockfile.

| Library | Version | Deps | Own size | Installed |
| --- | --- | --- | --- | --- |
| **openjsxl** | `1.1.0` | 1 | 33 KB | 0.60 MB |
| **ExcelJS** | `4.4.0` | 192 | 20.8 MB | 50.5 MB |
| **SheetJS** | `0.20.3` | 0 | 7.7 MB | 7.8 MB |

_openjsxl: Footprint of the exact published version; timing uses the workspace bundle identified above. Its sole runtime dependency is its own @openjsxl/core._

_ExcelJS: Measured with the benchmark's uuid 11.1.1 override; other transitives resolve during this dated clean install._

_SheetJS: Official distribution: https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz._

openjsxl ships **zero third-party runtime dependencies** (both its packages are pure TypeScript over
platform Web APIs). The dependency counts above describe the measured distributions; SheetJS
0.20.3 also declares no runtime dependencies. *Installed* includes the library itself.


## Read

Parse a real `.xlsx` and materialize every cell value.

#### numbers

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 12 ms · 70.4 MB | 44 ms · 94.4 MB | 25 ms · 77.0 MB |
| `100k` | 107 ms · 82.8 MB | 218 ms · 275.8 MB | 232 ms · 136.0 MB |
| `1M` | 858 ms · 209.0 MB | 1.65 s · 800.9 MB | 2.64 s · 527.8 MB |

#### strings

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 19 ms · 72.2 MB | 60 ms · 96.8 MB | 32 ms · 79.8 MB |
| `100k` | 107 ms · 83.6 MB | 220 ms · 205.5 MB | 220 ms · 136.2 MB |
| `1M` | 927 ms · 211.7 MB | 1.75 s · 714.3 MB | 2.25 s · 535.8 MB |

#### styled

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 13 ms · 71.0 MB | 45 ms · 94.8 MB | 30 ms · 79.4 MB |
| `100k` | 103 ms · 86.1 MB | 214 ms · 209.0 MB | 262 ms · 137.9 MB |
| `1M` | 929 ms · 229.9 MB | 1.91 s · 799.8 MB | 2.85 s · 545.0 MB |


## Write

Serialize N cells to `.xlsx` bytes. openjsxl's *streamed* column is fed a lazy row source (a
generator) — the honest streaming case, where the full row array never exists in memory.

#### numbers

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 9.7 ms · 64.5 MB | 9.5 ms · 62.5 MB | 58 ms · 106.2 MB | 16 ms · 76.1 MB |
| `100k` | 77 ms · 94.9 MB | 77 ms · 70.4 MB | 363 ms · 322.9 MB | 141 ms · 137.0 MB |
| `1M` | 741 ms · 395.4 MB | 760 ms · 100.3 MB | 3.26 s · 1.24 GB | 2.66 s · 526.8 MB |

#### strings

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 9.8 ms · 65.9 MB | 9.9 ms · 62.8 MB | 68 ms · 117.7 MB | 17 ms · 75.5 MB |
| `100k` | 88 ms · 119.5 MB | 87 ms · 67.5 MB | 380 ms · 345.0 MB | 170 ms · 143.2 MB |
| `1M` | 857 ms · 501.0 MB | 880 ms · 96.6 MB | 3.47 s · 1.56 GB | 2.39 s · 552.8 MB |

#### styled

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 23 ms · 68.2 MB | 23 ms · 64.5 MB | 85 ms · 121.1 MB | — |
| `100k` | 209 ms · 120.4 MB | 209 ms · 79.6 MB | 560 ms · 349.3 MB | — |
| `1M` | 2.03 s · 555.4 MB | 2.03 s · 128.4 MB | 5.23 s · 1.72 GB | — |

### Output file size

#### numbers

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 35 KB | 35 KB | 39 KB | 92 KB |
| `100k` | 322 KB | 322 KB | 334 KB | 875 KB |
| `1M` | 3.3 MB | 3.3 MB | 3.4 MB | 8.7 MB |

#### strings

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 59 KB | 59 KB | 66 KB | 132 KB |
| `100k` | 574 KB | 574 KB | 410 KB | 1.2 MB |
| `1M` | 5.7 MB | 5.7 MB | 3.7 MB | 12.2 MB |

#### styled

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 43 KB | 43 KB | 49 KB | — |
| `100k` | 404 KB | 404 KB | 422 KB | — |
| `1M` | 4.1 MB | 4.1 MB | 4.2 MB | — |


## Read by format

The same `numbers` data authored in four container formats, each read by the libraries that support
it and materialized cell-by-cell. xlsx reuses the ExcelJS-authored fixture; `.xlsb`/`.ods` are
authored by **SheetJS** (a real producer that writes those containers — the same author-then-read
shape ExcelJS→xlsx already uses); `.csv` is written directly (plain text, no producer). ExcelJS reads
only xlsx/csv. Python reference rows, when available below, use the exact same fixture hashes.
Format compatibility is measured with the recorded producer version; old producer results are not reused.

#### xlsx

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 12 ms · 70.2 MB | 62 ms · 94.4 MB | 29 ms · 76.8 MB |
| `100k` | 97 ms · 83.4 MB | 213 ms · 279.5 MB | 233 ms · 135.3 MB |
| `1M` | 867 ms · 208.9 MB | 1.66 s · 810.1 MB | 2.57 s · 513.4 MB |

#### xlsb

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 16 ms · 70.1 MB | 15 ms · 71.5 MB |
| `100k` | 64 ms · 122.2 MB | 148 ms · 116.5 MB |
| `1M` | 748 ms · 461.4 MB | 1.77 s · 415.7 MB |

#### ods

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 21 ms · 74.5 MB | 35 ms · 83.6 MB |
| `100k` | 141 ms · 163.8 MB | 320 ms · 183.5 MB |
| `1M` | 1.54 s · 804.2 MB | 3.29 s · 853.8 MB |

#### csv

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 4.7 ms · 64.5 MB | 13 ms · 89.7 MB | 7.0 ms · 68.0 MB |
| `100k` | 43 ms · 114.1 MB | 74 ms · 136.2 MB | 56 ms · 109.9 MB |
| `1M` | 593 ms · 406.7 MB | 667 ms · 595.3 MB | 1.14 s · 458.5 MB |


## Failed or unsupported measurements

These cells are shown as — above; their timing and memory measurements are excluded.

| Library | Operation | Workload | Format | Cells | Reason |
| --- | --- | --- | --- | --- | --- |
| SheetJS | write | styled | xlsx | 10k | unsupported: SheetJS community build does not emit cell styles |
| SheetJS | write | styled | xlsx | 100k | unsupported: SheetJS community build does not emit cell styles |
| SheetJS | write | styled | xlsx | 1M | unsupported: SheetJS community build does not emit cell styles |


## Reference: Python

No matching Python reference measurements. Legacy or different-fixture results are omitted.
Reference numbers for Python readers/writers are gathered out-of-band (they need a
Python environment, not npm). Run the companion script to populate this section:

```sh
python3 packages/bench/py/bench_py.py   # writes packages/bench/.cache/python.json
pnpm bench --render-only                 # preserve the JS measurement date and versions
```


## Methodology

- **Isolation.** Every `(library, operation, workload, size)` cell runs in its own
  `node --expose-gc` child process, so a cell's memory and timing reflect that library alone.
- **Timing.** `performance.now()` around the operation; the reported value is the **median** of the
  measured iterations (2–5 depending on size), each preceded by a full GC, after warmup runs that
  are discarded.
- **Memory.** Peak `process.memoryUsage().rss` of the **first, cold run only** (which also serves as
  JIT warmup) — sampled on a 3 ms timer that fires during async awaits, plus a post-op read for
  synchronous ops. Measuring across warmed iterations would over-report: V8 doesn't return freed heap
  to the OS, so later iterations carry earlier ones' residue. The cold figure is what a single call
  costs; it is absolute, so it includes loading the library.
- **Equal work.** Every writer is handed the identical dataset; every reader parses the identical
  ExcelJS-authored file and materializes every cell into JS values. Outside worker timing, the
  harness compares each returned read checksum with the deterministic workload's expected sum,
  requiring a finite value within max(1e-9, 32 × Number.EPSILON × |expected|). Mismatches are excluded
  and listed above; this checksum is a sanity check, not proof of complete semantic equivalence.
  SheetJS writes with `compression: true` so its output is deflated like the
  others'.
- **SheetJS styled write** is shown as —: emitting cell styles is a SheetJS Pro feature, absent from this community build, so measuring it would compare unequal work.
- **ExcelJS streaming.** ExcelJS also ships a streaming API (`stream.xlsx`); it is not benchmarked
  here. The ExcelJS columns use its standard in-memory Workbook API.
- **Not measured here.** CI-run benchmarks (too noisy) and internal micro-benchmarks. Python numbers
  are gathered out-of-band with a companion script and clearly marked as reference.

Regenerate with `pnpm bench` (add `--quick` for a fast smoke run). See
[`packages/bench`](../packages/bench) for the harness.
