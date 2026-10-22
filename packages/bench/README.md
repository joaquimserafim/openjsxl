# @openjsxl/bench

Private benchmark harness for openjsxl. **Never published** — this is where the competitor
libraries (ExcelJS, SheetJS) live as devDependencies, so the shipped packages (`openjsxl`,
`@openjsxl/core`) stay zero-dependency. The output is [`docs/benchmarks.md`](../../docs/benchmarks.md).

## Run

```sh
pnpm bench            # full matrix (10k / 100k / 1M cells) → docs/benchmarks.md
pnpm bench --quick    # 10k only, 1 iteration — a fast smoke test of the harness itself
pnpm --filter @openjsxl/bench sizes  # exact-artifact size matrix → .cache/sizes.json
```

`pnpm bench` builds `openjsxl` first (so it benchmarks the built workspace bundle),
authors controlled synthetic datasets with the comparator writers (cached under `.cache/`,
gitignored), then runs every cell and rewrites the report. Fixture sidecars record the producer's
version/source, locked artifact integrity, dependency lock hash, dataset dimensions/workload,
recipe revision, and file hash; a mismatch regenerates the
fixture. These are reproducibility checks for this private harness, not a security boundary for
untrusted uploads or direct worker invocation. Use this harness only with its generated datasets.

Raw results record the measurement date, machine, installed versions and sources, lockfile and
built-bundle hashes, and fixture hashes. `--render-only` preserves that snapshot; legacy JS results
without it must be remeasured. Quick/partial runs replace the cached matrix and report, so rerun the
full matrix before publishing claims. Older Python/size caches are omitted until refreshed with
matching provenance.

Fresh million-cell XLSB fixture creation can take several minutes and several GiB of memory in
SheetJS. This authoring step is outside the timed readers; matching fixture caches avoid repeating it.

Flags: `--sizes 10k,100k`, `--workloads numbers,strings,styled`, `--ops read,write`, `--iters N`,
`--warmup N`.

## What it measures

- **Libraries:** openjsxl (buffered `writeXlsx` + streamed `streamXlsx`), ExcelJS `4.4.0`,
  SheetJS `0.20.3` from the [official distribution](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz).
  ExcelJS uses the reviewed `uuid@11.1.1` override; the result metadata records the installed value.
- **Workloads:** 10-column sheets of 10k / 100k / 1M cells, in three flavours — numbers-heavy,
  strings-heavy (realistic repetition, so shared-string tables matter), and styled.
- **Operations:** read (parse → materialize every cell) and write (serialize N cells → bytes).
- **Read by format:** the same `numbers` data as `.xlsx` / `.xlsb` / `.ods` / `.csv`, each read by
  the libraries that support it. xlsx reuses the ExcelJS fixture; `.xlsb`/`.ods` are authored by
  SheetJS (a real producer that writes those containers); `.csv` is written directly. `src/formats.mjs`.
- **Library size:** `node packages/bench/src/sizes.mjs` packs each exact artifact and installs the
  local tarball with `npm install --omit=dev --ignore-scripts`. It uses the same official SheetJS
  distribution and ExcelJS uuid override, verifies comparator tarball bytes and npm metadata
  against the reviewed lockfile integrity, and
  records its transitive dependency count and on-disk footprint (`src/sizes.mjs` → `.cache/sizes.json`,
  folded into the report). The openjsxl footprint uses the exact published version, while timing
  uses the identified workspace bundle; there is no floating `latest` lookup. Other transitive
  dependencies resolve during the dated clean install. SheetJS 0.20.3 has no runtime dependencies.
- **Metrics:** median wall-time and peak process RSS.

## Why the numbers are trustworthy

- **Process isolation.** Every `(library, op, workload, size)` cell runs in its own
  `node --expose-gc` child (`worker.mjs`), spawned strictly one at a time. No library's heap or CPU
  contends with another's — the single biggest source of bogus benchmark numbers.
- **Warmup + median.** Warmup iterations (JIT, module init) are discarded; the reported time is the
  median of the measured ones. The heap is GC'd before each measured iteration.
- **Equal work, enforced.** One dataset feeds every writer; one ExcelJS-authored fixture feeds every
  reader (realistic shared strings + a styles table); readers materialize every cell into a checksum
  sink so nothing is optimized away. SheetJS writes with `compression: true` so file sizes compare.
- **Honest gaps.** SheetJS styled-write is reported as `—`, not a fast fake number — emitting cell
  styles is a SheetJS Pro feature absent from the community build, so measuring it would compare unequal
  work. openjsxl's streamed writer is fed a *lazy* row source, the real streaming use case.

## Python reference (out-of-band)

openpyxl and python-calamine numbers are gathered separately (they need Python, not npm) and merged
into the report as clearly-labelled reference rows:

```sh
python3 -m venv .venv && . .venv/bin/activate      # Python 3.9–3.14
pip install -r packages/bench/py/requirements.txt
python3 packages/bench/py/bench_py.py  # validates the current JS run's exact fixture hashes
pnpm bench --render-only             # merge references without relabelling the JS measurement
```

Run these commands from the repository root after a full JS matrix. Python results carry the JS
measurement identifier and fixture hashes, plus their own date and package versions; a later JS run
requires new Python measurements. Size results must match the recorded source/version, override,
artifact integrity, and dependency-lock snapshot before the report includes them.

## Layout

```
src/
  workloads.mjs   deterministic datasets (numbers / strings / styled; materialized + lazy)
  adapters/       one file per library: read / write against a shared interface
  worker.mjs      runs ONE isolated cell (warmup + timed iters + peak-RSS sampling)
  run.mjs         orchestrator: author fixtures, spawn workers serially, collect
  report.mjs      results → docs/benchmarks.md
  machine.mjs     hardware/runtime stamp   stats.mjs  median + human formatting
py/bench_py.py    openpyxl + python-calamine reference (subprocess-isolated, same fixtures)
```
