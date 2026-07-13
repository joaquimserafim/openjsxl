# @openjsxl/fuzz

Private fuzzing harness for openjsxl. **Never published** — [`fast-check`](https://github.com/dubzzz/fast-check)
lives here as the only devDependency, so the shipped packages (`openjsxl`, `@openjsxl/core`) stay
zero-dependency. Two complementary halves hammer the two directions the library can fail.

## Half A — writer property fuzzing (fast-check)

Generated workbook trees are fed to `writeXlsx`. Two corpora ([`src/arbitraries.ts`](src/arbitraries.ts)):
a **guaranteed-valid** corpus (unique names, finite numbers, XML-safe strings, legal styles, valid
tables/DV/CF — reliably writes) and a **hostile** corpus (non-plain prototypes, unknown keys,
transparent Proxies, poisoned values: `NaN`/`±Infinity`, control chars, invalid `Date`s, invalid
enums). The split matters: an all-hostile corpus rejects ~99.8% of inputs *before* producing bytes, so
the resolve/re-read and determinism checks would be near-vacuous. Four invariants
([`src/writer-properties.ts`](src/writer-properties.ts)):

- **P1 resolve-or-typed-error** (valid ∪ hostile) — `writeXlsx` either resolves to bytes `openXlsx`
  can re-read, or throws `XlsxError('invalid-input')`. A `TypeError`/`RangeError`/other code is a defect.
- **P4 valid round-trips** (valid) — a legitimately-valid workbook always writes AND re-opens (the
  writer never rejects valid input); this is where the resolve+re-read arm fires at ~100%.
- **P2 deterministic bytes** (valid) — identical input → byte-identical output.
- **P3 scalar round-trip** (scalars) — every written scalar reads back unchanged.

## Half B — corpus mutation fuzzing (seeded, zero-dep)

A seeded [xorshift](src/prng.ts) engine ([`src/mutate.ts`](src/mutate.ts)) corrupts every committed
fixture — generic byte flips/zero-runs/truncation/duplication, **zip-structure** pokes (EOCD /
central-directory / local-header counts, offsets, sizes, CRCs, name lengths), and **XML-ish**
blow-ups where markers sit in the clear — then replays each mutant against all four openers plus
format detection ([`src/replay.ts`](src/replay.ts)). The invariant every reader must uphold on **any**
bytes: **resolve, or throw a typed `XlsxError`** — never a `TypeError`/`RangeError`/bare throw. A
crasher is reproducible from `{fixture, integer seed}`.

## Run

```sh
pnpm vitest run packages/fuzz                    # the CI smoke suites (fixed seed, <1s)
pnpm --filter @openjsxl/fuzz fuzz                # the long local run (gated, budgeted)
```

The long run is tuned by env vars:

| var | default | meaning |
| --- | --- | --- |
| `FUZZ_MS` | `60000` | wall-clock budget for the mutation campaign |
| `FUZZ_RUNS` | `5000` | fast-check `numRuns` per writer property |
| `FUZZ_RSS_MB` | `2048` | RSS ceiling — the campaign aborts if exceeded (leak/OOM guard) |
| `FUZZ_SEED` | `42` | seed for fast-check and the mutation base (change to explore) |

```sh
FUZZ_MS=120000 FUZZ_SEED=7 pnpm --filter @openjsxl/fuzz fuzz
```

## CI posture

The smoke suites are ordinary vitest files under [`src/__tests__/`](src/__tests__), so the root gate
(`pnpm vitest run`) covers them with **zero new plumbing** — a fixed-seed fast-check pass (`numRuns`
≈ 60) plus a few hundred fixed-seed mutants, in well under a second, with **no wall-clock assertions**
(CI machines vary). The smoke also asserts the mutation engine actually *bites* (a nonzero typed-reject
tally), so a no-op fuzzer can't pass silently. The budgeted long run is gated behind `FUZZ_LONG=1` and
never runs in CI.

## Triage → promote

A long-run crasher is minimized (chunk-removal binary search, [`src/triage.ts`](src/triage.ts)) and
its reproducer written to the gitignored `crashers/` dir (raw + minimized bytes + a trace). A confirmed
crasher is then fixed and **promoted by hand** to `packages/fixtures/data/edge-*` with a verbatim-read
regression (the checklist at the bottom of [`../fixtures/data/README.md`](../fixtures/data/README.md)) —
never auto-committed.
