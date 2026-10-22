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

**Source snapshot:** lockfile SHA-256 `a09142b230a845d99f8090a7463f6404955da7d3c66fd7a9d62dcbb90e95b76e`; core bundle SHA-256
`b4fdabca4d56ba2567e6923e0e111965a2221254d00614a6de90ddd962f6da89`. The raw result cache also records locked comparator artifact
integrities and fixture hashes.

## Library size

A clean production install (`npm install --omit=dev --ignore-scripts`) of each exact artifact — measured 2026-10-21 on
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
| `10k` | 11 ms · 69.5 MB | 46 ms · 88.5 MB | 26 ms · 76.8 MB |
| `100k` | 90 ms · 82.5 MB | 191 ms · 221.5 MB | 232 ms · 136.4 MB |
| `1M` | 799 ms · 208.1 MB | 1.63 s · 815.6 MB | 2.58 s · 539.1 MB |

#### strings

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 17 ms · 71.2 MB | 62 ms · 91.3 MB | 31 ms · 79.1 MB |
| `100k` | 105 ms · 82.9 MB | 213 ms · 201.3 MB | 217 ms · 135.6 MB |
| `1M` | 909 ms · 211.5 MB | 1.74 s · 675.9 MB | 2.22 s · 540.8 MB |

#### styled

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 11 ms · 70.0 MB | 45 ms · 89.8 MB | 28 ms · 78.7 MB |
| `100k` | 96 ms · 85.7 MB | 224 ms · 276.7 MB | 259 ms · 137.2 MB |
| `1M` | 859 ms · 227.1 MB | 1.81 s · 795.9 MB | 2.74 s · 527.5 MB |


## Write

Serialize N cells to `.xlsx` bytes. openjsxl's *streamed* column is fed a lazy row source (a
generator) — the honest streaming case, where the full row array never exists in memory.

#### numbers

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 9.2 ms · 65.4 MB | 8.4 ms · 61.6 MB | 59 ms · 94.3 MB | 18 ms · 75.3 MB |
| `100k` | 73 ms · 93.7 MB | 72 ms · 69.5 MB | 350 ms · 250.8 MB | 150 ms · 135.7 MB |
| `1M` | 748 ms · 393.6 MB | 839 ms · 101.4 MB | 3.37 s · 1.31 GB | 2.64 s · 531.5 MB |

#### strings

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 10 ms · 66.0 MB | 9.3 ms · 62.2 MB | 66 ms · 105.8 MB | 20 ms · 75.4 MB |
| `100k` | 88 ms · 115.6 MB | 84 ms · 66.7 MB | 381 ms · 328.7 MB | 150 ms · 142.1 MB |
| `1M` | 838 ms · 502.1 MB | 927 ms · 96.3 MB | 3.51 s · 1.58 GB | 2.50 s · 552.0 MB |

#### styled

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 19 ms · 67.0 MB | 18 ms · 63.1 MB | 86 ms · 116.6 MB | — |
| `100k` | 161 ms · 126.4 MB | 157 ms · 78.2 MB | 563 ms · 298.4 MB | — |
| `1M` | 1.57 s · 555.6 MB | 1.56 s · 125.8 MB | 5.43 s · 1.76 GB | — |

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
| `10k` | 11 ms · 69.2 MB | 46 ms · 88.5 MB | 25 ms · 76.1 MB |
| `100k` | 92 ms · 82.1 MB | 214 ms · 268.8 MB | 242 ms · 132.9 MB |
| `1M` | 783 ms · 206.5 MB | 1.66 s · 805.1 MB | 2.56 s · 535.7 MB |

#### xlsb

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 10 ms · 70.0 MB | 15 ms · 70.6 MB |
| `100k` | 66 ms · 126.6 MB | 152 ms · 120.6 MB |
| `1M` | 949 ms · 513.0 MB | 1.68 s · 456.8 MB |

#### ods

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 17 ms · 73.8 MB | 35 ms · 82.8 MB |
| `100k` | 135 ms · 164.9 MB | 354 ms · 181.3 MB |
| `1M` | 1.49 s · 804.7 MB | 3.25 s · 984.2 MB |

#### csv

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 5.3 ms · 64.1 MB | 12 ms · 83.7 MB | 5.9 ms · 67.8 MB |
| `100k` | 52 ms · 115.4 MB | 67 ms · 132.7 MB | 47 ms · 107.3 MB |
| `1M` | 627 ms · 477.6 MB | 654 ms · 569.1 MB | 1.04 s · 460.8 MB |


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
