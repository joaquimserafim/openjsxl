# openjsxl 1.1.0 against 1.0.0

_A one-off release record, measured 2026-11-03. `docs/benchmarks.md` holds the full 1.1.0 matrix;
this page records the 1.0.0 comparison behind the 1.1.0 changelog's costs line._

**Method.** Two full `pnpm bench` runs, one after the other, on one machine (Apple M2 Pro · 12
cores · 32 GiB · darwin arm64 · Node v24.22.1), each library cell in its own isolated process. Both
runs use the same harness, build toolchain and lockfile, and the same 18 cached fixtures (identical
hashes). The 1.1.0 run is the release tree. The 1.0.0 run swaps in `v1.0.0`'s `packages/core/src`
and `packages/openjsxl/src`, built with the same toolchain, so it measures 1.0.0's source, not its
published bundle. The load average stayed between about 3 and 4 during both runs; an earlier pair,
measured under a load average near 20, was discarded.

**Comparability.** The comparator libraries are identical in both runs. At 1M cells their times
moved between the runs by -4.8% to +12.0%, median -0.1%.

Median wall-time · peak RSS at **1M cells**, openjsxl only:

| Operation | 1.0.0 | 1.1.0 | Time |
| --- | --- | --- | --- |
| Read `.xlsx`, numbers | 711 ms · 208.7 MB | 858 ms · 209.0 MB | +21% |
| Read `.xlsx`, strings | 807 ms · 212.1 MB | 927 ms · 211.7 MB | +15% |
| Read `.xlsx`, styled | 788 ms · 231.3 MB | 929 ms · 229.9 MB | +18% |
| Read `.xlsx` (format lane), numbers | 712 ms · 213.2 MB | 867 ms · 208.9 MB | +22% |
| Write, numbers | 747 ms · 394.3 MB | 741 ms · 395.4 MB | -1% |
| Write, strings | 816 ms · 504.3 MB | 857 ms · 501.0 MB | +5% |
| Write, styled | 1.50 s · 555.8 MB | 2.03 s · 555.4 MB | +35% |
| Stream write, numbers | 761 ms · 99.9 MB | 760 ms · 100.3 MB | 0% |
| Stream write, strings | 847 ms · 96.4 MB | 880 ms · 96.6 MB | +4% |
| Stream write, styled | 1.54 s · 126.0 MB | 2.03 s · 128.4 MB | +32% |
| Read `.ods` | 1.40 s · 800.5 MB | 1.54 s · 804.2 MB | +10% |
| Read `.csv` | 549 ms · 477.9 MB | 593 ms · 406.7 MB | +8% |
| Read `.xlsb` | — | 748 ms · 461.4 MB | — |

1.0.0 shows — for `.xlsb` because its reads fail the harness's read checksum: it misses the
compact cell records SheetJS 0.20.3 writes, which 1.1.0 decodes. Smaller sizes are noisier: the same
`.xlsx` read, measured in two lanes of one run, differs by up to about 12% at 10k and 100k cells,
against about 1% at 1M.

Raw result caches (SHA-256):

- 1.0.0: `c73ccd23319dea8cd3c31449b9da63dcd5dd56c1ece77b118dc1f696b5ef9327`
- 1.1.0: `b4d0f3a21769379786a7a40c83168c88b48742443ab7137a1835f936b054a399`
