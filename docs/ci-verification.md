# CI verification — F11.29

Reviewed **2026-10-23**, starting from `1be06cf`, on Node **24.21.0** and pnpm **11.28.7**. This is
a local simulation of `.github/workflows/ci.yml` in a fresh clone. It does not replace the hosted
CI/CodeQL verification tracked by F11.18.

## Problem

`packages/bench/src/__tests__/provenance.test.ts` calls `librarySnapshot()`, which located the
workspace facade's `package.json` by resolving `openjsxl` through its exports map. That resolution
lands on `dist/index.js`, which exists only after a build. The test never loads the bundle, but a
checkout without a prior build failed `pnpm test`: hosted CI, which ran Test before Build, and the
README's install → check → typecheck → test development flow. Local runs passed only because ignored
`packages/*/dist` builds existed.

The blocking production audit also preceded every code check. A failed audit skipped Lint, Typecheck,
Test and Build, so an advisory or an audit-endpoint outage hid their results.

## Change

- **Provenance lookup.** `librarySnapshot()` reads the workspace facade's metadata from its source
  manifest, `packages/openjsxl/package.json`. The comparators (ExcelJS, SheetJS) keep the installed
  package lookup. The reported version is unchanged: the old lookup resolved the same file.
- **CI step order.** Test stays before Build, so CI keeps enforcing that tests run from source with
  no build. The job name `lint · typecheck · test · build` still matches the order.
- **CI step conditions.** The install step has `id: install`. Lint, Typecheck, Test and Build run
  under `if: ${{ !cancelled() && steps.install.outcome == 'success' }}`.
  - A failed audit still fails the job, but no longer skips the code checks.
  - A failed code check no longer skips the checks after it.
  - A failed install, cancellation or timeout still skips them.
- **Audit policy unchanged.** The production audit blocks; the full audit is triage only.

## Clean-clone simulation

The clone was made with `git clone --no-local` at `1be06cf`, with the changed `ci.yml` and
`provenance.mjs` copied in. It contains no `dist` and no local-only fixtures. The steps ran in
`ci.yml` order.

| Step | Command | Exit | Result |
| --- | --- | --- | --- |
| Install | `pnpm install --frozen-lockfile --offline` | 0 | Lockfile unchanged |
| Workspace audit | `pnpm audit` | 0 | No known vulnerabilities |
| Production audit | `pnpm audit --prod` | 0 | No known vulnerabilities |
| Lint | `pnpm check` | 0 | 309 files, 4 existing nonblocking warnings |
| Typecheck | `pnpm typecheck` | 0 | — |
| Test (no `dist`) | `pnpm test` | 0 | 2,068 passed / 3 skipped / 2,071 collected |
| Build | `pnpm build` | 0 | — |

**Control.** Restoring `1be06cf`'s `provenance.mjs` in the same clone and removing `dist` makes
`pnpm test` exit 1, with 1 failed / 2,067 passed / 3 skipped / 2,071 collected and
`Cannot find module '<clone>/packages/bench/node_modules/openjsxl/dist/index.js'`.

Afterwards, the clone's working tree showed only the two copied files. The build wrote only ignored
`dist` output.

**Skips.** The 3 skips are environment-gated:
- the 2 fuzz long-run tests, which run only when `FUZZ_LONG` is `1` or `true`;
- 1 comments test, which needs the local-only `packages/fixtures/local/SimpleWithComments.xlsx`.

A checkout that has that fixture skips 2.

## Not covered locally

- **Workflow lint.** actionlint was not available for this record.
- **Hosted runner and action behavior**, including the step-condition semantics. Per GitHub's
  expression documentation, a status-check function such as `cancelled()` in an `if` replaces the
  implicit `success()` condition.
- **Online install and caching.**
- **CodeQL and Dependabot runs.**

All of these remain part of F11.18's hosted verification.
