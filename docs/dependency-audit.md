# Dependency advisory review — F11.17

Reviewed **2026-10-21** against baseline `a8dc07f`, Node 24.21.0. Remediation owner for
this review: **Codex, implementing the owner's F11.17 authorization**. Future dependency
maintenance remains a repository maintenance task. This is a dated assessment, not a
promise that future audits will remain clean.

The initial full workspace audit reported **27 findings: 2 critical, 16 high, 8 moderate,
1 low** (21 distinct GHSA IDs). The production audit reported **3 high**, all in pnpm's
bootstrap dependency graph. They were tooling findings, but they still required remediation.
Neither public package ships these dependencies: `@openjsxl/core` has no dependencies and
`openjsxl` depends only on the matching first-party core package.

No advisory is suppressed or retained under risk acceptance. The table covers every initial
finding, including both brace-expansion branches and both Vitest/mocker entries. Upstream
fixed ranges are the ranges reported by the dated registry audit; the installed resolution
column records the chosen remediation.

## Dependency paths and reachability

Path aliases below expand to **all** paths from the initial audit (`root` is the workspace
root; `bench` is the private `@openjsxl/bench` package):

| Path | Full initial dependency path(s) |
| --- | --- |
| S | bench → xlsx |
| U | bench → exceljs → uuid |
| E | root → tsup → esbuild; root → tsup → bundle-require → esbuild |
| B1 | bench → exceljs → archiver → archiver-utils → glob → minimatch → brace-expansion; bench → exceljs → archiver → zip-stream → archiver-utils → glob → minimatch → brace-expansion; bench → exceljs → unzipper → fstream → rimraf → glob → minimatch → brace-expansion |
| B2 | bench → exceljs → archiver → readdir-glob → minimatch → brace-expansion |
| P | root → tsup → postcss; root → tsup → postcss-load-config → postcss; root → vitest → vite → postcss; root → vitest → @vitest/mocker → vite → postcss; root → vitest → vite-node → vite → postcss |
| M | packageManagerDependencies → pnpm |
| V | root → vitest; root → vitest → @vitest/mocker |
| D | root → vitest → vite; root → vitest → @vitest/mocker → vite; root → vitest → vite-node → vite |
| T | root → vitest → tinypool |

| Package/path | Affected command and input reachability |
| --- | --- |
| xlsx (S) | `pnpm bench`: crafted workbook parser input; normal runs generate local synthetic fixtures. |
| uuid (U) | `pnpm bench`: ExcelJS x14 conditional-format IDs call `v4()` without arguments; the vulnerable buffer-taking v3/v5/v6 APIs are not used. |
| esbuild (E) | `pnpm build`, Vitest config loading: repository uses build/transform APIs, not the vulnerable Windows development server. |
| brace-expansion (B1/B2) | `pnpm bench`: ExcelJS archive/glob dependencies accept patterns; synthetic buffer benchmarks do not accept caller-supplied glob patterns. |
| postcss (P) | `pnpm build` / `pnpm test`: hostile CSS sourceMappingURL input; current library entry points are TypeScript and do not load untrusted CSS. |
| nanoid (P → nanoid) | PostCSS build/test dependency; vulnerable custom-generator zero-size calls are not made by project code. |
| source-map-js (P → source-map-js) | Build/test tooling parsing indexed source maps with hostile section offsets. |
| pnpm (M) | `pnpm install`: untrusted workspace configuration, lock dependency paths, or tarball manifest names, respectively. |
| vitest (V) | `pnpm test` / watch: hostile redirect-mock path through the test development server. Tests/configuration are repository inputs. |
| vite (D) | `pnpm test` / watch: development-server editor launch endpoint and hostile filename/editor arguments; no standalone Vite server is configured. |
| tinypool (T) | `pnpm test`: worker options become code execution gadgets after a separate prototype-pollution primitive. |

## Advisory dispositions

Every row has owner **Codex / F11.17** and review date **2026-10-21**. “Updated” means the
vulnerable version is removed from the resolved graph, even where normal repository use
did not exercise the vulnerable API.

| Advisory | Package / prior version | Severity | Path | Upstream fixed range | Disposition / installed resolution |
| --- | --- | --- | --- | --- | --- |
| [GHSA-4r6h-8v6p-xvw6](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) | xlsx 0.18.5 | high | S | `>=0.19.3` | Updated: 0.20.3 (official integrity-pinned distribution) |
| [GHSA-5pgg-2g8v-p4x9](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9) | xlsx 0.18.5 | high | S | `>=0.20.2` | Updated: 0.20.3 (official integrity-pinned distribution) |
| [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq) | uuid 8.3.2 | moderate | U | `>=11.1.1` | Updated: 11.1.1 (scoped ExcelJS override) |
| [GHSA-g7r4-m6w7-qqqr](https://github.com/advisories/GHSA-g7r4-m6w7-qqqr) | esbuild 0.27.7 | low | E | `>=0.28.1` | Updated: 0.28.2 under tsup; 0.28.1 under Vite |
| [GHSA-mh99-v99m-4gvg](https://github.com/advisories/GHSA-mh99-v99m-4gvg) | brace-expansion 1.1.16 | high | B1 | `>=1.1.17` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-mh99-v99m-4gvg](https://github.com/advisories/GHSA-mh99-v99m-4gvg) | brace-expansion 2.1.2 | high | B2 | `>=2.1.3` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-fxqj-rqcc-2cmp](https://github.com/advisories/GHSA-fxqj-rqcc-2cmp) | postcss 8.5.18 | moderate | P | `>=8.5.23` | Updated: 8.5.31 (compatible lock refresh) |
| [GHSA-rgw5-rvv9-x895](https://github.com/advisories/GHSA-rgw5-rvv9-x895) | brace-expansion 2.1.2 | high | B2 | `>=2.1.4` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-rgw5-rvv9-x895](https://github.com/advisories/GHSA-rgw5-rvv9-x895) | brace-expansion 1.1.16 | high | B1 | `>=1.1.18` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-2v37-7h3g-55p8](https://github.com/advisories/GHSA-2v37-7h3g-55p8) | nanoid 3.3.16 | high | P → nanoid | `>=3.3.18` | Updated: 3.3.20 (compatible lock refresh) |
| [GHSA-vx52-2968-3vc6](https://github.com/advisories/GHSA-vx52-2968-3vc6) | pnpm 11.9.0 | high | M | `>=11.11.0` | Updated: 11.28.6 (package-manager pin and bootstrap lock) |
| [GHSA-c59q-g84q-2gj5](https://github.com/advisories/GHSA-c59q-g84q-2gj5) | pnpm 11.9.0 | high | M | `>=11.11.0` | Updated: 11.28.6 (package-manager pin and bootstrap lock) |
| [GHSA-vq4v-j7r6-jq4m](https://github.com/advisories/GHSA-vq4v-j7r6-jq4m) | pnpm 11.9.0 | high | M | `>=11.11.0` | Updated: 11.28.6 (package-manager pin and bootstrap lock) |
| [GHSA-82fw-gwwq-j7x9](https://github.com/advisories/GHSA-82fw-gwwq-j7x9) | vitest 3.2.7 | moderate | V (vitest) | `>=4.1.11` | Updated: 4.1.12 (reviewed major migration) |
| [GHSA-82fw-gwwq-j7x9](https://github.com/advisories/GHSA-82fw-gwwq-j7x9) | @vitest/mocker 3.2.7 | moderate | V (mocker) | `>=4.1.11` | Updated: 4.1.12 (reviewed major migration) |
| [GHSA-q2hr-2g5m-vwhr](https://github.com/advisories/GHSA-q2hr-2g5m-vwhr) | brace-expansion 1.1.16 | moderate | B1 | `>=1.1.21` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-q2hr-2g5m-vwhr](https://github.com/advisories/GHSA-q2hr-2g5m-vwhr) | brace-expansion 2.1.2 | moderate | B2 | `>=2.1.7` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-qhr7-859c-m2p7](https://github.com/advisories/GHSA-qhr7-859c-m2p7) | brace-expansion 1.1.16 | high | B1 | `>=1.1.20` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-qhr7-859c-m2p7](https://github.com/advisories/GHSA-qhr7-859c-m2p7) | brace-expansion 2.1.2 | high | B2 | `>=2.1.6` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-6j4f-fj2g-mc7p](https://github.com/advisories/GHSA-6j4f-fj2g-mc7p) | brace-expansion 1.1.16 | high | B1 | `>=1.1.19` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-6j4f-fj2g-mc7p](https://github.com/advisories/GHSA-6j4f-fj2g-mc7p) | brace-expansion 2.1.2 | high | B2 | `>=2.1.5` | Updated: 1.1.21 and 2.1.7 (compatible lock refresh) |
| [GHSA-68fv-2mgg-jv7q](https://github.com/advisories/GHSA-68fv-2mgg-jv7q) | source-map-js 1.2.1 | high | P → source-map-js | `>=1.2.2` | Updated: 1.2.2 (compatible lock refresh) |
| [GHSA-5gmw-xhrv-c9v3](https://github.com/advisories/GHSA-5gmw-xhrv-c9v3) | tinypool 1.1.1 | critical | T | `>=2.1.1` | Removed: Removed by Vitest 4 migration |
| [GHSA-85c8-ppgw-ccpr](https://github.com/advisories/GHSA-85c8-ppgw-ccpr) | tinypool 1.1.1 | critical | T | `>=2.1.2` | Removed: Removed by Vitest 4 migration |
| [GHSA-698v-4r3x-3vcv](https://github.com/advisories/GHSA-698v-4r3x-3vcv) | vite 7.3.6 | moderate | D | `>=7.3.7` | Updated: 7.3.7 (explicit compatible selection) |
| [GHSA-r825-w8mm-5pxw](https://github.com/advisories/GHSA-r825-w8mm-5pxw) | vite 7.3.6 | moderate | D | `>=7.3.7` | Updated: 7.3.7 (explicit compatible selection) |
| [GHSA-587h-7wmj-h5pm](https://github.com/advisories/GHSA-587h-7wmj-h5pm) | vite 7.3.6 | high | D | `>=7.3.7` | Updated: 7.3.7 (explicit compatible selection) |

## Reviewed dependency choices

- **pnpm 11.28.6:** stays on the existing major and its Node >=22.13 requirement. The
  repository already requires Node >=24. Both the manifest pin and the package-manager
  section of the lockfile must agree; merely changing `packageManager` leaves stale audit
  entries in an existing multi-document lockfile. The matching `devEngines.packageManager`
  declaration opts pnpm 11 into refreshing that bootstrap graph; both pins are exact.
  `pnpm install --lockfile-only --ignore-scripts --no-frozen-lockfile` regenerated it.
  No global package-manager installation was changed.
- **Vitest 4.1.12 / Vite 7.3.7:** Vitest 3 has no patched release for this inventory.
  The [v4 migration guide](https://v4.vitest.dev/guide/migration.html) and
  [4.1.12 release](https://github.com/vitest-dev/vitest/releases/tag/v4.1.12) were reviewed.
  Vitest 4 removes the tinypool dependency. The repository uses none of the removed pool,
  browser, or third-argument test-options APIs. The reporter/config hooks used by the
  collected-test floor remain available; CLI failure modes are checked separately from
  the full suite. Vite 7 is explicit to avoid an unrelated Vite 8 migration.
- **tsup 8.5.1 → esbuild 0.28.2 override:** scoped to this exact parent release because
  tsup's declared `^0.27.0` range has no fix. This is a reviewed override outside the parent
  range, not an upstream-supported tsup dependency combination. The repository uses the
  JS build API for ESM TypeScript bundles, without esbuild serving. The 0.28 changes preserve
  that API; builds, public declarations, examples and packed consumers verify compatibility.
  `bundle-require`'s `esbuild >=0.18` peer accepts the chosen version. Remove this override
  when a reviewed tsup release declares a patched esbuild range. tsup 9 is deferred because
  its newer Node floor is unrelated to this fix.
- **ExcelJS 4.4.0 → uuid 11.1.1 override:** scoped to this exact parent release, also outside
  its declared range. ExcelJS's only uuid imports are CommonJS `v4`, called with no arguments
  for x14 conditional-format identifiers. Version 11.1.1 retains that API and CommonJS;
  later ESM-only majors are unnecessary. Regression tests exercise both call sites and a
  public workbook write/read round-trip. Remove the override when ExcelJS declares a
  patched compatible uuid range or no longer uses uuid.
- **PostCSS, nanoid, source-map-js, brace-expansion:** refresh within existing compatible
  ranges; both brace-expansion majors are patched without global overrides.

## SheetJS source and license

The npm `xlsx@0.18.5` comparator is replaced with **SheetJS CE 0.20.3** from the upstream
[Node installation instructions](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/):
[`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz).
The manifest records that exact URL; pnpm records its SHA-512 integrity:

```text
sha512-oLDq3jw7AcLqKWH2AhCpVTZl8mf6X2YReP+Neh0SJUzV/BdZYjth94tG5toiMB1PPrYtxOCfaoUCkvtuH+3AJA==
```

The reviewed artifact is 2,409,319 bytes, SHA-256
`8dc73fc3b00203e72d176e85b50938627c7b086e607c682e8d3c22c02bb99fe8`.
Its 26 regular files contain Apache-2.0 licenses at `package/LICENSE` and
`package/dist/LICENSE`, no NOTICE file, no dependencies, no install/prepare lifecycle hooks,
and no traversal paths or symlinks. The dependency is not vendored or included in public
package tarballs. The integrity pin detects changes to the fetched artifact; it is not a
claim that arbitrary upstream code is safe.

The adapter's date handling and buffer read/write behavior are checked against ExcelJS.
The complete benchmark matrix is rerun because comparator behavior and timing can change
across this upgrade. Fixtures are regenerated when producer provenance changes; reports
retain the measurement's recorded versions, date and machine instead of relabeling cached
results. Legacy Python/size results lacking matching provenance are omitted.

The benchmark remains private development tooling for locally generated synthetic fixtures.
It is not an upload service or a security sandbox. Do not feed untrusted workbooks to its
worker CLI. See [the benchmark runbook](../packages/bench/README.md).

## Validation record

- Both `pnpm audit --prod --json` and `pnpm audit --json` exit **0**, with **zero advisories**
  at every severity, including the pnpm bootstrap graph.
- A clean temporary workspace install using `pnpm install --frozen-lockfile` exits **0**;
  the lockfile is byte-identical afterwards. The allowed esbuild installation hooks succeed.
- Both public builds, root/workspace TypeScript checks, and all 14 runnable examples pass.
  Core and facade public `index.d.ts` files remain byte-identical to the baseline.
- Six CLI guard probes preserve direct focused runs, enforce the root collection floor even
  with a reporter override, reject empty discovery, and retain exactly one GitHub Actions
  reporter/error annotation where expected.
- Public tarballs contain only manifests, README, LICENSE and built distribution files.
  A temporary consumer resolves both locally packed packages, exercises buffered/streaming
  write and readback, and loads main/formula entry points through ESM and Node CommonJS.
  Its runtime graph contains only `openjsxl` and `@openjsxl/core`.

- Final Biome check passes (281 files); the guarded root suite passes **1,683 tests**, with
  **2 skipped / 1,685 collected**. This includes 22 comparator/provenance/checksum tests.
- Independent finder/verifier review confirmed two harness defects: missing size-artifact
  integrity enforcement and missing dataset dimensions in fixture cache identity. Both are
  fixed. A changed tarball, forged metadata, and mismatched supplemental hashes now reject;
  changing row count regenerates the fixture. No confirmed F11.17 harness finding remains open; the separate parser gap is recorded below.
- Exact-artifact size measurements pass, including verification of the downloaded SheetJS
  bytes against the lockfile SRI before installation. Install scripts are disabled.

The first complete matrix exposed a separate, pre-existing compatibility gap: openjsxl ignores undocumented XLSB
short cell variants emitted by SheetJS 0.20.3, so it reads only the first column of these fixtures.
The benchmark now checks read sums against its deterministic dataset and excludes mismatched
measurements. The XLSB parser correction is separately scoped as F11.28 in the local release plan;
no dependency risk acceptance is involved.

The final full matrix exits **0** and records **93 cells: 87 valid measurements, 3 unsupported
SheetJS styled writes, and 3 excluded openjsxl XLSB checksum failures**. Every other read checksum
matches its recipe; all recorded fixture and lockfile hashes match the bytes used. Root/facade
README claims and [the full report](benchmarks.md) use these dated measurements. Old Python
reference rows are omitted because they lack matching provenance. The dependency work is complete;
the XLSB compatibility correction remains open as separate release work.
