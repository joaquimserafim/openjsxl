# CLAUDE.md — how to work on openjsxl

Operating contract for agents (and humans) working in this repo. The *what/why* is
[ROADMAP.md](./ROADMAP.md); the scoped tracker is [IMPLEMENTATION.md](./IMPLEMENTATION.md);
the living session checkpoint is `progress.md` (gitignored — read it first, update it last).

## The prime rule

**Nothing gets implemented before it is defined, scoped, and broken into tasks in
IMPLEMENTATION.md.** Work one feature (`Fx.y`) at a time, in milestone order, and get the
owner's explicit go-ahead ("proceed") before starting a feature.

## Working agreements (owner-set; do not bend)

1. **Every commit needs prior approval**: present a report + the proposed commit message,
   then WAIT for the owner's "commit". Never commit or push on your own.
2. **No** `Co-Authored-By` / "generated with" trailers in commit messages.
3. Tests live in `__tests__/` folders beside the code they cover. Always.
4. **Version bumps only at the owner's explicit request** — and READMEs + `examples/` must
   be updated *before* any bump. The owner runs `git push`, `pnpm -r publish --access
   public` (never the npm CLI), and tags. See PUBLISHING.md.
5. Adversarial review before committing correctness-sensitive code (parsers, emitters,
   validators). Fix confirmed findings + pin regressions first.
6. `progress.md` is never committed; keep it current so a cold session can resume.

## Quality gates (all must pass before a feature is "done")

```sh
pnpm biome check .          # judge by EXIT CODE — never pipe through tail/grep (a past
                            # regression shipped because a pipe hid the failure)
pnpm -r exec tsc --noEmit
pnpm vitest run             # from the repo root (per-package vitest finds no tests)
```

Plus, for any change touching the **writer**:

- **Byte-identity**: input that uses no new feature must produce the exact pre-feature
  bytes. In-tree golden pins cover this in CI; for belt-and-braces, build the pre-styles
  writer (`git worktree add /tmp/old-writer 1839e71`, `pnpm install && pnpm build` in its
  `packages/core`) and byte-compare `writeXlsx` outputs on: no-date, date,
  multi-sheet+sparse, 1904, empty-sheet.
- **Corpus property** (writer/__tests__/bridge-styles.test.ts): every readable fixture
  round-trips losslessly OR fails with a TYPED `XlsxError('invalid-input')` — never a bare
  throw, never silent change. Extend the snapshot when you add carried state.
- **openpyxl cross-validation** both directions (independent implementation): read our
  output with `warnings.simplefilter("error")`, and read a real openpyxl-authored fixture
  with our reader. Use a throwaway venv (openpyxl 3.1.5) in the scratchpad, never in-repo.

## Core invariants (violating any of these is a bug)

- **Zero runtime dependencies** in `@openjsxl/core` and `openjsxl`. Platform Web APIs only
  (`DecompressionStream`/`CompressionStream`, `TextEncoder`/`TextDecoder`). Dev/bench
  tooling is exempt (private packages only).
- **Reader degrades, writer rejects — with SHARED bounds.** The tolerant reader clamps or
  drops out-of-bounds producer values; the writer refuses invalid input with a typed
  error; both sides read the SAME single-sourced constants (`ooxml/a1.ts`: MAX_ROW,
  MAX_COL, MAX_COL_WIDTH, MAX_ROW_HEIGHT; `ooxml/styles.ts`: HEX_COLOR, MAX_COLOR_INDEX,
  MAX_INDENT; `utils/chars.ts`: isXmlSafe). Whatever the reader can return, the bridge can
  write — or the write fails typed.
- **One shared model.** What a reader accessor returns IS what the writer accepts
  (styles, geometry, merges, hyperlinks, state). No parallel "writer flavors" of a type.
- **Single-read validation (TOCTOU).** Read each caller-supplied property exactly once
  into a local before validating/emitting — getters/Proxies must not be able to change a
  value between validation and emission. Validate plain-object prototypes
  (`isPlainRecord`), reject unknown keys, and `escapeAttr`/`isXmlSafe` everything emitted.
- **Deterministic bytes.** Identical input → identical output. No timestamps, no
  randomness in the writer.
- **Schema element order** is load-bearing (worksheet: dimension → sheetViews → cols →
  sheetData → mergeCells → hyperlinks → … → legacyDrawing; workbook: workbookPr →
  bookViews → sheets). Excel repair-prompts on order violations.
- **Unused features emit nothing.** New writer capabilities must be invisible (empty
  string, absent part, absent attribute) when unused — that is what keeps byte-identity.
- **Performance is adversarial-input-safe.** No O(n²) on attacker-controlled counts, no
  absurd array lengths from hostile refs (see the F4.4 grid-cap and F4.6 merge-sweep
  precedents).

## Adversarial review protocol

- Spawn finders with distinct lenses (spec-conformance, round-trip/bridge, hostile-input,
  algorithm-correctness), then one refuting verifier per finding. Track
  CONFIRMED / REFUTED / **UNVERIFIED** — a dead verifier (rate limit) is UNVERIFIED, not
  refuted; re-verify those claims yourself, empirically, with probe scripts.
- Probes live in the session scratchpad, **never in the repo**. After any review round,
  check `git status` for stray `tmp-*.test.ts` / probe files and delete them — dead agents
  have leaked files into `__tests__/` before, breaking the suite.

## Fixtures

- Committed corpus: `packages/fixtures/data/` — programmatic (regenerate with
  `pnpm fixtures`; byte-deterministic) + real-producer files (openpyxl/Excel/LibreOffice;
  provenance in `data/README.md`, vendored-file licenses in `THIRD_PARTY.md`).
- Differently-licensed real files go in gitignored `packages/fixtures/local/` only.
- New format behavior ⇒ new fixture + a test that reads it verbatim; follow the checklist
  at the bottom of `data/README.md`.

## Repo map (orientation)

```
packages/core/src/
  zip/      reader: EOCD walk + inflate-on-demand      writer/zip.ts: deflate + headers
  xml/      SAX tokenizer (never a DOM), chunk-safe streaming variant
  ooxml/    rels graph, workbook, shared strings, styles table, a1 refs, dates, bounds
  reader/   openXlsx / streamSheetRows / Workbook / Worksheet accessors (lazy, cached)
  writer/   writeXlsx (types→sheet→styles→workbook→zip), from-workbook.ts = the bridge
  utils/    isXmlSafe + small shared helpers
packages/openjsxl/     facade: re-exports core (users install this)
packages/fixtures/     private corpus + generator
examples/              runnable consumer-style examples (kept working; CI-adjacent docs)
```

Public API grows only through `packages/core/src/index.ts` (facade re-exports `*`).
