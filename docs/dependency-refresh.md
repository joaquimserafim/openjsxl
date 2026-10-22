# Release dependency refresh — F11.19

Reviewed **2026-10-22**, starting from `41a1227`, on Node **24.21.0**. This is the
dependency snapshot prepared for 1.0.1; the package versions remain 1.0.0 until an
explicit release step. It does not close the remaining code, documentation, Excel,
or hosted CI/CodeQL/Dependabot release prerequisites.

The [F11.17 advisory dispositions](dependency-audit.md) remain the security baseline.
No advisory suppression, new runtime dependency, package-source substitution, or
unreviewed major migration is part of this refresh.

## Direct dependency decisions

Registry versions and upstream releases were checked on the review date. Public
core/facade runtime requirements stay at Node >=24. All changes below are private
development or package-manager dependencies.

| Dependency | Before → selected | Decision and evidence |
| --- | --- | --- |
| pnpm | 11.28.6 → 11.28.7 | [Stable patch release](https://github.com/pnpm/pnpm/releases/tag/v11.28.7), retaining Node >=22.13. Both manifest pins and the bootstrap lock document agree. The `latest-11` registry tag still pointed to 11.28.6; the exact 11.28.7 release is not a prerelease. |
| Biome | 2.5.3 → 2.6.2 | [Compatible minor](https://github.com/biomejs/biome/releases/tag/@biomejs%2Fbiome@2.6.2), with formatter and recommended-rule changes. See the scope treatment below. |
| `@types/node` | 24.13.3 → 24.19.2 | Latest reviewed Node 24 declarations; minimum TypeScript 5.6 remains compatible with the selected compiler. |
| TypeScript | 5.9.3 retained | Latest 5.x. [Version 7](https://devblogs.microsoft.com/typescript/announcing-typescript-7-0/) changes the compiler implementation/API; the current declaration bundler consumes the TypeScript JavaScript API. Defer that migration. |
| tsup | 8.5.1 retained | Latest 8.x. [Version 9](https://github.com/egoist/tsup/releases/tag/v9.0.0) raises the Node 24 floor to 24.11 and changes build dependencies; requires a separate compatibility decision. |
| Vite | 7.3.7 retained | Latest 7.x. [Version 8](https://vite.dev/guide/migration) changes the transform/bundler stack to Oxc/Rolldown; defer the migration. |
| Vitest | 4.1.12 retained | Latest 4.x. [Version 5](https://vitest.dev/guide/migration/) changes mock defaults, filtering and reporters; preserve the reviewed test semantics. |
| fast-check | 3.23.2 retained | Latest 3.x. [Version 4](https://fast-check.dev/blog/2025/03/10/whats-new-in-fast-check-4-0-0/) changes generated values/defaults and RNG behavior; seed reproducibility needs a separate migration. |
| ExcelJS | 4.4.0 retained | Latest stable release. The 4.4.1 prerelease is excluded. Its compatible transitive graph is refreshed below. |
| SheetJS | official CE 0.20.3 retained | The [upstream distribution](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/) remains the reviewed URL and integrity pin. npm's older `xlsx` package is not a replacement. |

Also deferred: pnpm 12's package-manager migration and Node 26 type declarations.
The scoped `tsup@8.5.1 > esbuild@0.28.2` and `exceljs@4.4.0 > uuid@11.1.1`
overrides retain their F11.17 compatibility evidence and removal criteria. Later uuid
majors remove CommonJS, used by ExcelJS. No blanket override was added.

The fixtures workspace has no registry dependencies. The fuzz workspace retains
fast-check and its dependency graph. Public core has no runtime dependencies;
the facade and examples retain only their first-party workspace edges. Python
reference benchmark pins are unchanged and their historical results remain omitted
until independently remeasured against the current fixtures.

## Compatible transitive refresh

`pnpm update -r --depth Infinity --lockfile-only --ignore-scripts` refreshed the full
workspace within declared ranges. The committed lockfile records every resolution
and integrity; its two documents include the package-manager bootstrap graph.

| Workspace/path | Changes |
| --- | --- |
| Package-manager bootstrap | pnpm and `@pnpm/exe` plus its platform packages 11.28.6 → 11.28.7; reflink/detect-libc unchanged. |
| Root formatter | Biome and platform binaries 2.5.3 → 2.6.2, including new optional Android binaries. |
| Root Node declarations | `@types/node` 24.13.3 → 24.19.2; `undici-types` 7.18.2 → 7.24.6. ExcelJS's separate Node 14 types remain unchanged. |
| Root build/test tools | Rollup 4.62.2 → 4.64.12 and its native packages; acorn 8.17.0 → 8.19.0; sourcemap-codec 1.5.5 → 1.6.0; picomatch 4.0.5 → 4.0.7; lilconfig 3.1.3 → 3.1.4. Vite's esbuild 0.28.1 unifies with the existing 0.28.2. |
| Private ExcelJS comparator | dayjs 1.11.21 → 1.11.23; JSZip 3.10.1 → 3.10.2; traverse 0.3.9 → 0.3.10. The traverse patch adds 92 helper package instances; the combined ExcelJS/SheetJS closure grows from 86 to 178 instances. |
| Other workspaces | No registry resolution changes in fixtures, fuzz, core, facade or examples. |

Rollup's new optional `@napi-rs/lzma-linux-x64-gnu@1.5.1` is an upstream
[cross-build helper](https://github.com/rollup/rollup/pull/6467). The exact release
source and artifact agree; shipped Rollup JavaScript does not import it. It has no
lifecycle scripts and is constrained to Linux x64 glibc. Its higher optional Node
requirement does not change the public packages' runtime requirement.

The [traverse patch](https://github.com/ljharb/js-traverse/releases/tag/v0.3.10) is
reached through ExcelJS → unzipper → binary → chainsaw. Its source matches the exact
upstream tag. All 95 introduced comparator package versions (three replacements
plus 92 additions) were checked against lockfile SHA-512 values, archive paths,
licenses and lifecycle scripts: no path escapes/symlinks or installation hooks;
MIT is available for each, including JSZip's existing dual-license choice.
The reviewed tooling archives likewise match registry and lockfile integrity.
Only the already-allowed esbuild installation hook runs during the clean install.
These checks establish provenance and packaging properties, not reproducible native
binaries or a claim that dependency code cannot contain defects.

The five existing deprecated ExcelJS transitives (fstream, glob, inflight,
lodash.isequal and rimraf) remain private comparator dependencies. A compatible
refresh cannot remove them; no current registry advisory is suppressed. The harness
continues to accept only its controlled synthetic fixtures.

## Formatter compatibility

Biome 2.6 enables Markdown and YAML handling by default. Their formatter/linter
settings are explicitly disabled, and Markdown embedded-code analysis remains off,
preserving the previous check scope and pnpm's own lockfile formatting. JavaScript
and TypeScript checks remain enabled with the recommended preset.

The new formatter reflows 18 existing parameterized test files. Their parsed
TypeScript syntax trees are identical before and after formatting. No runtime source
was changed. Four new `useMathMinMax` style warnings remain visible; applying the
suggested unsafe runtime rewrites is outside this dependency refresh. The Biome
command exits successfully without suppressing that rule.

## Validation and freeze

Validation on Node 24.21.0 / pnpm 11.28.7:

- Clean frozen install in a new workspace and store passes, including esbuild's
  allowed hook; the lockfile is byte-identical afterwards.
- Production and full audits both exit 0 with zero advisories at every severity.
- Biome exits 0 (304 files, four style warnings described above); root and recursive
  TypeScript checks pass. The guarded suite passes **1,756 tests / 2 skipped /
  1,758 collected**, including the corpus bridge and fixed-seed fuzz smoke.
- Seven CLI probes verify focused runs, the root collection floor despite a reporter
  override, empty-discovery failure and exactly one GitHub Actions reporter.
- Both builds and all 14 examples pass. The main core bundle and public main-entry
  declarations remain byte-identical to the F11.18 baseline. Prepack reproduces all
  current runtime/declaration files exactly.
- Empty-cache offline consumers install only the local public tarballs: core alone,
  or facade plus core. Packed facade dependencies rewrite to exact core 1.0.0.
  Contents are limited to manifests, licenses, READMEs and distribution files;
  main/formula imports and CommonJS loads, strict NodeNext declarations, entry-point
  isolation, both writers and current XLSB bridges pass. Independent openpyxl
  read/write checks pass with warnings treated as errors.
- Comparator baseline/current controls pass buffered workbooks, nested chainsaw
  calls, unzipper extraction and an ordered ZIP's streaming read. An ExcelJS-authored
  ZIP triggers the same [upstream streaming issue](https://github.com/exceljs/exceljs/issues/3064)
  in both graphs; it is not introduced here and the benchmark uses buffered APIs.

Extended fuzz passes on seeds **42** and **20261022**: four writer properties × 5,000
cases × two seeds = **40,000 property cases**. Recorded mutation runs cover all 55
fixtures and produce **1,494,350 opener outcomes with zero crashes**:

| Seed | Mutated inputs | Inputs whose bytes changed | Peak sampled descendant RSS |
| --- | --- | --- | --- |
| 42 | 157,080 | 156,423 | 1,327,513,600 bytes |
| 20261022 | 141,790 | 141,300 | 1,447,084,032 bytes |

Each mutation campaign used `FUZZ_MS=60000` and `FUZZ_RSS_MB=2048`; both stopped on
the time budget, not RSS. An external supervisor also enforced 900 seconds and
3 GiB aggregate descendant RSS per command, sampled every 0.2 seconds. The seed 42
full run passed in 81.442 seconds; its mutation-only repeat took 62.101 seconds to
recover counters hidden by successful-test console capture. The other full run took
181.030 seconds. The table counts only the recorded mutation repeat for seed 42,
without double-counting its earlier run.

The full isolated benchmark was rerun after the dependency change, with freshly
validated fixture provenance: **93 cells, 90 valid measurements, 3 expected
unsupported SheetJS styled writes**. All **57 read checksums match exactly**, including
XLSB at every size. Exact-artifact installed sizes were remeasured. The
[full report](benchmarks.md) and root/facade README figures use this snapshot:
1M-cell numeric XLSX read **805 ms**, buffered write **729 ms**, streamed write
**754 ms**, and XLSB read **990 ms**. These are measurements of the unchanged
workspace runtime with refreshed tooling/comparators, not a runtime speedup claim.

Final local candidate packs include the updated README. Compared with the tested
packs, only the facade README changed; executable files, declarations, manifests
and licenses are identical. A fresh offline install contains only the two public
packages, totaling **528,180 logical file bytes** (filesystem allocation differs).

| Local pack | Tarball bytes | Logical installed bytes | SHA-256 |
| --- | --- | --- | --- |
| core 1.0.0 | 121,903 | 499,279 | `3672439b146e2ce414e6c460bb71c7c42d0fb3feca4a0727243817129b13c07f` |
| facade 1.0.0 | 11,298 | 28,901 | `b100bffb07eda1d36cf1cd8d3dd220ac7268d8eeb988d7768de3532e64e42856` |

These packs contain unreleased workspace changes under the unchanged version;
they are validation artifacts, not the published 1.0.0 tarballs.

Dependency inputs, SHA-256 (paths relative to the repository):

```text
3ed43fd527e99c5b7f8f8ebbdd6ad99ce355ab5f19b752551dad01322148d31c  package.json
0900ab1e27020ea6de1bda912bf847d6b3ce4042d5db196d4320e19d8b14300b  pnpm-workspace.yaml
d55cf0f92fcff742726deb66f80e0587234264b1d8dbdf4867b84139a5c59457  pnpm-lock.yaml
384dc06d518ccbf1a99a8f810ba801200dcb513587bacc6ea0bd29a0e7edde83  packages/bench/package.json
bf759c200a331b006e847f1762fa91b7cce62927ec0997e6473fca05d30bc53a  packages/core/package.json
9a5f21884fee87dc72b739c13686f0952a1fa35a995a970e052e9c17c0a2d450  packages/fixtures/package.json
3214fb8bb8207c805557b4cdba16d36d5e1afa237710dcf1195e76e04ea42b52  packages/fuzz/package.json
cce17e2052b751ad00f69fb6ee4a88e868fd494210e96a5bf2f7df09dd045fb9  packages/openjsxl/package.json
f903f2a4ec04b3cd66d448ceacafc00a01ddf340ae03d7738b8fd224dc1aa8f3  examples/package.json
```

Any dependency-input change invalidates this snapshot and requires F11.19 validation
again. Later source changes require fresh release gates and artifact validation;
benchmark-input changes require new measurements. A version bump also changes
manifests and packed dependency rewriting, so final release packages must be checked
again. This record does not authorize a commit, push, version bump or publication.
