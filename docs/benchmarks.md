# openjsxl benchmarks

_Measured by `pnpm bench` on 2026-10-21. Re-rendering preserves this measurement date._

**Machine:** Apple M2 Pro · 12 cores · 32 GiB · darwin arm64 · Node v24.21.0

Each number is the **median wall-time** (over warmed iterations) and the **peak process RSS of a
single cold call** (library load + one operation), measured in a **fresh, isolated Node process per
cell** so no library colours another's timing or memory. Workloads are 10 columns wide;
`10k`/`100k`/`1M` count cells. Read fixtures are authored once by ExcelJS (realistic shared
strings + a styles table) and read by all three JS readers. Full methodology and the fairness
caveats are at the bottom.

Libraries: **openjsxl** `1.0.0` (`workspace`); **ExcelJS** `4.4.0` (`4.4.0`; override uuid 11.1.1); **SheetJS** `0.20.3` ([official distribution](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz)). Peak RSS is absolute and includes loading the library.

**Source snapshot:** lockfile SHA-256 `a09142b230a845d99f8090a7463f6404955da7d3c66fd7a9d62dcbb90e95b76e`; core bundle SHA-256
`5e753d112a40440898cb2ceec7769c0e3e3e63fc2b28fe6e8e8b94e05a7852b6`. The raw result cache also records locked comparator artifact
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
| `10k` | 11 ms · 68.4 MB | 42 ms · 88.3 MB | 24 ms · 75.7 MB |
| `100k` | 93 ms · 82.8 MB | 191 ms · 193.3 MB | 225 ms · 134.9 MB |
| `1M` | 769 ms · 208.0 MB | 1.56 s · 812.3 MB | 2.49 s · 542.6 MB |

#### strings

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 16 ms · 71.8 MB | 60 ms · 91.2 MB | 30 ms · 78.1 MB |
| `100k` | 101 ms · 83.9 MB | 198 ms · 196.1 MB | 213 ms · 135.9 MB |
| `1M` | 838 ms · 210.0 MB | 1.67 s · 787.5 MB | 2.11 s · 539.3 MB |

#### styled

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 26 ms · 70.7 MB | 47 ms · 90.1 MB | 28 ms · 79.2 MB |
| `100k` | 99 ms · 84.7 MB | 224 ms · 258.2 MB | 248 ms · 138.3 MB |
| `1M` | 832 ms · 228.3 MB | 1.74 s · 835.8 MB | 2.75 s · 552.2 MB |


## Write

Serialize N cells to `.xlsx` bytes. openjsxl's *streamed* column is fed a lazy row source (a
generator) — the honest streaming case, where the full row array never exists in memory.

#### numbers

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 9.2 ms · 64.6 MB | 8.0 ms · 62.1 MB | 56 ms · 94.4 MB | 17 ms · 73.8 MB |
| `100k` | 72 ms · 94.2 MB | 67 ms · 68.9 MB | 335 ms · 251.9 MB | 126 ms · 135.8 MB |
| `1M` | 686 ms · 390.7 MB | 668 ms · 100.6 MB | 3.16 s · 1.43 GB | 2.37 s · 523.9 MB |

#### strings

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 19 ms · 66.8 MB | 18 ms · 62.6 MB | 94 ms · 110.1 MB | 26 ms · 76.0 MB |
| `100k` | 81 ms · 115.3 MB | 80 ms · 66.7 MB | 357 ms · 327.7 MB | 140 ms · 142.9 MB |
| `1M` | 808 ms · 502.2 MB | 854 ms · 96.0 MB | 3.41 s · 1.59 GB | 2.24 s · 551.7 MB |

#### styled

| Cells | openjsxl (buffered) | openjsxl (streamed) | ExcelJS | SheetJS |
| --- | --- | --- | --- | --- |
| `10k` | 20 ms · 67.6 MB | 18 ms · 63.7 MB | 84 ms · 117.0 MB | — |
| `100k` | 152 ms · 120.7 MB | 143 ms · 79.2 MB | 515 ms · 297.9 MB | — |
| `1M` | 1.42 s · 555.6 MB | 1.42 s · 126.5 MB | 5.11 s · 1.76 GB | — |

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
| `10k` | 12 ms · 68.9 MB | 41 ms · 88.5 MB | 24 ms · 76.3 MB |
| `100k` | 88 ms · 81.7 MB | 183 ms · 193.1 MB | 343 ms · 134.1 MB |
| `1M` | 838 ms · 207.5 MB | 1.90 s · 740.8 MB | 2.84 s · 524.0 MB |

#### xlsb

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | — | 14 ms · 70.3 MB |
| `100k` | — | 144 ms · 120.7 MB |
| `1M` | — | 1.99 s · 456.8 MB |

#### ods

| Cells | openjsxl | SheetJS |
| --- | --- | --- |
| `10k` | 17 ms · 75.2 MB | 35 ms · 84.0 MB |
| `100k` | 143 ms · 163.2 MB | 316 ms · 184.4 MB |
| `1M` | 1.70 s · 805.3 MB | 3.24 s · 955.1 MB |

#### csv

| Cells | openjsxl | ExcelJS | SheetJS |
| --- | --- | --- | --- |
| `10k` | 4.9 ms · 64.3 MB | 11 ms · 83.3 MB | 5.7 ms · 68.0 MB |
| `100k` | 50 ms · 115.4 MB | 134 ms · 138.2 MB | 65 ms · 104.4 MB |
| `1M` | 612 ms · 477.4 MB | 654 ms · 568.8 MB | 1.06 s · 460.4 MB |


## Failed or unsupported measurements

These cells are shown as — above; their timing and memory measurements are excluded.

| Library | Operation | Workload | Format | Cells | Reason |
| --- | --- | --- | --- | --- | --- |
| SheetJS | write | styled | xlsx | 10k | unsupported: SheetJS community build does not emit cell styles |
| SheetJS | write | styled | xlsx | 100k | unsupported: SheetJS community build does not emit cell styles |
| SheetJS | write | styled | xlsx | 1M | unsupported: SheetJS community build does not emit cell styles |
| openjsxl | read | numbers | xlsb | 10k | Read checksum mismatch: expected 74995650.00000311, received 7492500 (tolerance 5.328761432110917e-7) |
| openjsxl | read | numbers | xlsb | 100k | Read checksum mismatch: expected 7499956500.00169, received 749925000 (tolerance 0.000053290396095929466) |
| openjsxl | read | numbers | xlsb | 1M | Read checksum mismatch: expected 749999565001.5581, received 74999250000 (tolerance 0.005329067427350922) |


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
