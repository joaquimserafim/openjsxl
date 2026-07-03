# @openjsxl/bench

Private benchmark harness for openjsxl. **Never published** — this is where the competitor
libraries (ExcelJS, SheetJS) live as devDependencies, so the shipped packages (`openjsxl`,
`@openjsxl/core`) stay zero-dependency. The output is [`docs/benchmarks.md`](../../docs/benchmarks.md).

## Run

```sh
pnpm bench            # full matrix (10k / 100k / 1M cells) → docs/benchmarks.md
pnpm bench --quick    # 10k only, 1 iteration — a fast smoke test of the harness itself
```

`pnpm bench` builds `openjsxl` first (so it benchmarks the real published bundle, not TS source),
authors the read fixtures once (cached under `.cache/`, gitignored), then runs every cell and
rewrites the report.

Flags: `--sizes 10k,100k`, `--workloads numbers,strings,styled`, `--ops read,write`, `--iters N`,
`--warmup N`.

## What it measures

- **Libraries:** openjsxl (buffered `writeXlsx` + streamed `streamXlsx`), ExcelJS `4.4.0`,
  SheetJS `xlsx@0.18.5` (the last npm-registry release).
- **Workloads:** 10-column sheets of 10k / 100k / 1M cells, in three flavours — numbers-heavy,
  strings-heavy (realistic repetition, so shared-string tables matter), and styled.
- **Operations:** read (parse → materialize every cell) and write (serialize N cells → bytes).
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
  styles is a SheetJS Pro feature absent from the npm build, so measuring it would compare unequal
  work. openjsxl's streamed writer is fed a *lazy* row source, the real streaming use case.

## Python reference (out-of-band)

openpyxl and python-calamine numbers are gathered separately (they need Python, not npm) and merged
into the report as clearly-labelled reference rows:

```sh
python3.13 -m venv .venv && . .venv/bin/activate   # Python 3.9–3.13 (see py/requirements.txt)
pip install -r py/requirements.txt
python3 py/bench_py.py         # reads the same .cache/ fixtures → .cache/python.json
node src/run.mjs --render-only  # merge the Python column in without re-running the JS matrix
```

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
