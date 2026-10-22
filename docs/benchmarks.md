# openjsxl benchmarks

_Measured by `pnpm bench` on 2026-10-22. Re-rendering preserves this measurement date._

**Machine:** Apple M2 Pro · 12 cores · 32 GiB · darwin arm64 · Node v24.21.0

Each number is the **median wall-time** (over warmed iterations) and the **peak process RSS of a
single cold call** (library load + one operation), measured in a **fresh, isolated Node process per
cell** so no library colours another's timing or memory. Workloads are 10 columns wide;
`10k`/`100k`/`1M` count cells. Read fixtures are authored once by ExcelJS (realistic shared
strings + a styles table) and read by all three JS readers. Full methodology and the fairness
caveats are at the bottom.

Libraries: **openjsxl** `1.0.0` (`workspace`); **ExcelJS** `4.4.0` (`4.4.0`; override uuid 11.1.1); **SheetJS** `0.20.3` ([official distribution](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz)). Peak RSS is absolute and includes loading the library.

**Source snapshot:** lockfile SHA-256 `d55cf0f92fcff742726deb66f80e0587234264b1d8dbdf4867b84139a5c59457`; core bundle SHA-256
`b4fdabca4d56ba2567e6923e0e111965a2221254d00614a6de90ddd962f6da89`. The raw result cache also records locked comparator artifact
integrities and fixture hashes.

## Library size

A clean production install (`npm install --omit=dev --ignore-scripts`) of each exact artifact — measured 2026-10-22 on
Node 24.21.0. **Deps** counts the runtime packages the library pulls in
(transitive, excluding the library itself); **Own size** is the package's own unpacked tarball;
**Installed** is the library plus every runtime dependency on disk.
Comparator tarball bytes and npm metadata must match the artifact integrity in the recorded lockfile.

| Library | Version | Deps | Own size | Installed |
| --- | --- | --- | --- | --- |
| **openjsxl** | `1.0.0` | 1 | 25 KB | 0.52 MB |
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
| `10k` | 11 ms · 68.9 MB | 43 ms · 92.9 MB | 25 ms · 77.2 MB |
| `100k` | 93 ms · 81.5 MB | 199 ms · 193.3 MB | 234 ms · 135.0 MB |
| `1M` | 805 ms · 208.1 MB | 1.69 s · 819.0 MB | 2.62 s · 528.1 MB |

#### strings

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 16 ms · 71.3 MB | 61 ms · 95.5 MB | 32 ms · 79.6 MB |
| `100k` | 103 ms · 83.1 MB | 222 ms · 199.1 MB | 233 ms · 136.6 MB |
| `1M` | 862 ms · 210.0 MB | 1.73 s · 787.0 MB | 2.21 s · 538.2 MB |

#### styled

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 12 ms · 69.0 MB | 43 ms · 93.5 MB | 28 ms · 79.6 MB |
| `100k` | 95 ms · 84.7 MB | 204 ms · 207.8 MB | 253 ms · 137.3 MB |
| `1M` | 855 ms · 228.7 MB | 1.85 s · 830.7 MB | 2.79 s · 556.3 MB |


## Write

Serialize N cells to `.xlsx` bytes. openjsxl's *streamed* column is fed a lazy row source (a
generator) — the honest streaming case, where the full row array never exists in memory.

#### numbers

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 9.2 ms · 64.6 MB | 8.3 ms · 61.5 MB | 60 ms · 105.1 MB | 18 ms · 75.1 MB |
| `100k` | 77 ms · 93.8 MB | 74 ms · 69.8 MB | 345 ms · 323.9 MB | 181 ms · 133.4 MB |
| `1M` | 729 ms · 369.3 MB | 754 ms · 100.2 MB | 3.36 s · 1.51 GB | 2.38 s · 524.0 MB |

#### strings

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 10 ms · 65.8 MB | 9.5 ms · 62.3 MB | 64 ms · 116.3 MB | 19 ms · 76.6 MB |
| `100k` | 90 ms · 115.4 MB | 89 ms · 67.2 MB | 387 ms · 247.6 MB | 145 ms · 142.3 MB |
| `1M` | 836 ms · 501.8 MB | 826 ms · 96.6 MB | 3.48 s · 1.56 GB | 2.57 s · 551.9 MB |

#### styled

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 19 ms · 67.0 MB | 18 ms · 64.2 MB | 189 ms · 120.2 MB | — |
| `100k` | 158 ms · 126.8 MB | 156 ms · 74.6 MB | 546 ms · 301.9 MB | — |
| `1M` | 1.51 s · 556.0 MB | 1.55 s · 125.3 MB | 5.30 s · 1.76 GB | — |

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
| `10k` | 12 ms · 68.7 MB | 46 ms · 93.4 MB | 25 ms · 78.4 MB |
| `100k` | 91 ms · 82.0 MB | 200 ms · 224.2 MB | 229 ms · 134.7 MB |
| `1M` | 831 ms · 207.6 MB | 1.66 s · 809.0 MB | 2.67 s · 531.3 MB |

#### xlsb

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 10 ms · 69.7 MB | 15 ms · 70.8 MB |
| `100k` | 70 ms · 126.4 MB | 155 ms · 120.6 MB |
| `1M` | 990 ms · 491.1 MB | 1.76 s · 457.9 MB |

#### ods

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 17 ms · 74.3 MB | 36 ms · 83.7 MB |
| `100k` | 135 ms · 161.7 MB | 319 ms · 183.2 MB |
| `1M` | 1.56 s · 823.2 MB | 3.29 s · 961.2 MB |

#### csv

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 5.5 ms · 64.6 MB | 12 ms · 89.7 MB | 6.0 ms · 68.3 MB |
| `100k` | 52 ms · 115.6 MB | 68 ms · 135.3 MB | 53 ms · 107.5 MB |
| `1M` | 655 ms · 477.3 MB | 662 ms · 593.4 MB | 1.08 s · 459.3 MB |


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
- **Not measured here.** CI-run benchmarks (too noisy) and internal micro-benchmarks. Python numbers
  are gathered out-of-band with a companion script and clearly marked as reference.

Regenerate with `pnpm bench` (add `--quick` for a fast smoke run). See
[`packages/bench`](../packages/bench) for the harness.
