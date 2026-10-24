# Security policy

## Reporting a vulnerability

Please use GitHub's [private vulnerability reporting form](https://github.com/joaquimserafim/openjsxl/security/advisories/new)
for security issues in openjsxl. Do not disclose exploit details or sensitive workbooks in public
issues or pull requests.

A useful report includes the affected package/version or commit, runtime/version, entry point and
options, expected versus observed behavior, and a small reproducible example. For resource-exhaustion
reports, include the input size and observed time/memory under your test limits. Remove personal
data, credentials and unrelated proprietary content from attachments.

Reports are handled on a best-effort basis. We aim to assess impact, discuss a fix and coordinate
public disclosure with the reporter. There is no guaranteed response or remediation deadline or LTS
commitment. Please coordinate disclosure timing through the private report.

## Supported versions

Security fixes target the latest published release of `openjsxl` and `@openjsxl/core`. Upgrade to
the latest patch release; backports to older versions are not guaranteed. Reports against the
development branch are welcome, but unreleased code is not a supported published release.

**Version note, 2026-10-22:** the latest published version is **1.0.0**. This branch contains work
toward **1.0.1**, which is not yet released; its unchanged package version does not make these
changes available in the published 1.0.0 artifacts. In particular, `ReadOptions.verifyChecksums`,
the exact ZIP-entry size checks, the ZIP entry-overlap check and the read-each-part-once rule
described below are unreleased 1.0.1 changes, absent from 1.0.0.

## Processing untrusted files

The readers accept spreadsheet bytes and may encounter malformed or deliberately expensive inputs.
Resource exhaustion and incorrect interpretation of workbook data are relevant security concerns.
Tolerant parsing and typed errors help applications handle failures; they do not establish a
sandbox or make every accepted workbook safe for every use.

For ZIP-based readers (`openXlsx`, `openXlsb`, `openOds`) and the XLSX `streamSheetRows` path,
`ReadOptions` provides two guards in both published 1.0.0 and this branch:

| Option | Default | What it limits |
| --- | --- | --- |
| `maxPartBytes` | 2 GiB | Uncompressed output of one accessed ZIP entry. |
| `maxCompressionRatio` | 300 | Inflation relative to that entry's compressed size, with an 8 MiB output floor. |

These limits are **per entry**, not an aggregate workbook, process-memory, CPU-time or cell-count
budget. Many entries can each fit their limit; parsed cells, strings, styles, images and other
metadata can use additional memory. The input byte buffer is already allocated before the reader
receives it. Streaming rows does not remove the input/archive and shared-metadata costs, and
retaining streamed rows defeats its memory benefit. CSV/TSV has no ZIP container: `openCsv` does
not use these guards and parses its input synchronously.

On this unreleased 1.0.1 branch, these readers also require every ZIP entry to occupy its own bytes.
Before any part is read, an archive is rejected with `XlsxError` code `corrupt-zip` when two
entries share bytes, or when an entry's data reaches the central directory. This stops many ZIP
entries from naming one small compressed stream so that it inflates once per entry. Above the ZIP
layer, opening a workbook with `openXlsx` or `openXlsb` reads each worksheet, sheet-relationships,
comments, table, shared-strings, styles and theme part at most once. A later sheet whose worksheet
part was already read is left out of `sheets`, and a comments or table part that was already read
is skipped. Pictures, which `Worksheet.images()` and so `workbookToInput` read on demand, are not
covered yet. None of this limits total decompressed output: each part read is still bounded only by
the per-entry limits above, so a workbook of many distinct large parts can still use a large amount
of memory.

Applications that accept untrusted uploads should:

- Bound input bytes before calling the library, and limit concurrent jobs and retained results.
- Set lower, finite positive `maxPartBytes` and `maxCompressionRatio` values appropriate to the
  accepted documents. `Infinity` disables a guard; do not disable limits for untrusted input.
- Process files in a killable process with externally enforced time and memory limits. A Promise
  timeout in the same event loop does not stop synchronous parsing or reclaim its allocations.
- Treat cell text, formulas, links and media as untrusted application data. Apply the destination's
  escaping, link and media policies before displaying, exporting or opening that data elsewhere.

The library does not execute VBA/macros, follow hyperlinks, fetch external workbook references,
or resolve network resources while parsing. Formula text and producer-cached results are read as
data. Rewriting a workbook is not a general sanitization step: supported formula text, hyperlinks
and images can survive, and another spreadsheet application may act on them when opened.

## Formula evaluation is explicit

Reading or writing a workbook does not recalculate its formulas. Evaluation is an explicit call
through `openjsxl/formula` or `@openjsxl/core/formula`. The built-in evaluator does not execute
formula text as JavaScript or fetch external references; unsupported external references yield
error values.

`EvaluateOptions.maxCellVisits` defaults to **5,000,000 formula-cell evaluations**. Set a lower,
finite positive integer for untrusted workloads. Parser/evaluator depth and argument guards also
exist, but these are not comprehensive time or memory quotas: constructing the workbook model,
parsing text and work inside an individual function are not all charged as cell visits. Keep
external process limits in place even with a small evaluation budget.

Functions registered through `EvaluateOptions.functions`, and injected `now`/`random` callbacks,
are **trusted application JavaScript**. They run with the host's privileges and can perform I/O,
allocate memory or fail to return. The evaluator's budget cannot preempt them. Do not register
code supplied by an untrusted workbook or user; the function registry is not a sandbox.

## Corruption detection and authenticity

On this unreleased 1.0.1 branch, each fully consumed ZIP entry must match its declared uncompressed
size. Setting `ReadOptions.verifyChecksums: true` additionally checks its central-directory
**CRC-32**; this option defaults to **false**. Cancelled streams make no whole-entry size/checksum
claim for partially consumed entries, and unopened parts have not been validated by reading
other parts. Published 1.0.0 does not provide this option or the exact-size reconciliation.

CRC-32 detects accidental corruption. It is not a cryptographic signature, does not authenticate
the sender, and does not prevent intentional modification: an attacker can change the data and
its CRC together. Establish file authenticity separately through a trusted transport, signature
or independently trusted digest when the application requires it.

## Runtime packages and development tooling

`@openjsxl/core` has no third-party runtime dependencies. `openjsxl` depends only on its own core
package. This reduces dependency exposure; it does not eliminate vulnerabilities in our code or
the JavaScript runtime/platform APIs.

The private benchmark, test, build and package-manager toolchains have their own dependency graphs
and execute code during development or CI. An advisory in those tools is not automatically a
vulnerability in the published runtime packages, but may affect contributors or release integrity.
Please identify the affected environment in a report. Reviewed tooling findings are tracked in
[the dependency audit](docs/dependency-audit.md); its dated results are not a guarantee about future
advisories.
