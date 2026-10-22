import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { FORMAT_READERS, FORMATS } from "./formats.mjs";
import { machineLine } from "./machine.mjs";
import { matchingPythonSnapshot, matchingSizeSnapshot, requireProvenance } from "./provenance.mjs";
import { validateReadResults } from "./read-validation.mjs";
import { fmtBytes, fmtMs } from "./stats.mjs";

// Renders the collected results into docs/benchmarks.md. Every published number comes from an
// actual run — this file only formats, it never invents or extrapolates a value.

const DOCS = fileURLToPath(new URL("../../../docs/benchmarks.md", import.meta.url));
const PYTHON_CACHE = fileURLToPath(new URL("../.cache/python.json", import.meta.url));
const SIZES_CACHE = fileURLToPath(new URL("../.cache/sizes.json", import.meta.url));

const READ_COLS = ["openjsxl", "ExcelJS", "SheetJS"];
const WRITE_COLS = ["openjsxl (buffered)", "openjsxl (streamed)", "ExcelJS", "SheetJS"];

function pick(results, op, workload, sizeKey, label) {
	return results.find(
		(r) => r.op === op && r.workload === workload && r.sizeKey === sizeKey && r.label === label,
	);
}

// A time·memory cell. Missing/failed → em dash; the failure table explains each rejected cell.
function cellTimeMem(res) {
	if (!res?.ok) return "—";
	return `${fmtMs(res.timeMs)} · ${fmtBytes(res.peakRss)}`;
}

function table(header, rows) {
	const head = `| ${header.join(" | ")} |`;
	const sep = `| ${header.map(() => "---").join(" | ")} |`;
	const body = rows.map((r) => `| ${r.join(" | ")} |`).join("\n");
	return `${head}\n${sep}\n${body}`;
}

function failureSection(results) {
	const failed = results.filter((result) => !result.ok);
	if (failed.length === 0) return "";
	const text = (value) =>
		String(value ?? "unknown")
			.replace(/\|/g, "\\|")
			.replace(/[\r\n]+/g, " ");
	const rows = failed.map((result) =>
		[
			result.label ?? result.lib,
			result.op,
			result.workload,
			result.format ?? "xlsx",
			result.sizeKey,
			result.reason,
		].map(text),
	);
	return `\n## Failed or unsupported measurements\n\nThese cells are shown as — above; their timing and memory measurements are excluded.\n\n${table(["Library", "Operation", "Workload", "Format", "Cells", "Reason"], rows)}\n`;
}

// The streamed column's op is write-stream; every other write-side label is a plain write; on the
// read side every label is a read.
function opFor(op, label) {
	if (op === "read") return "read";
	return label === "openjsxl (streamed)" ? "write-stream" : "write";
}

function opTable(results, op, sizes, workload, render) {
	const labels = op === "read" ? READ_COLS : WRITE_COLS;
	const rows = sizes.map((s) => [
		`\`${s.key}\``,
		...labels.map((label) => render(pick(results, opFor(op, label), workload, s.key, label))),
	]);
	return table(["Cells", ...labels], rows);
}

function section(results, op, sizes, workloads) {
	let md = "";
	for (const workload of workloads) {
		md += `\n#### ${workload}\n\n${opTable(results, op, sizes, workload, cellTimeMem)}\n`;
	}
	return md;
}

function sizeSection(results, sizes, workloads) {
	// Output .xlsx size for writes — surfaces the inline-strings (openjsxl) vs shared-strings
	// (ExcelJS/SheetJS) trade-off, and that styled files carry a styles table.
	let md = "";
	for (const workload of workloads) {
		const rows = sizes.map((s) => [
			`\`${s.key}\``,
			...WRITE_COLS.map((label) => {
				const res = pick(results, opFor("write", label), workload, s.key, label);
				return res?.ok && res.outBytes != null ? fmtBytes(res.outBytes) : "—";
			}),
		]);
		md += `\n#### ${workload}\n\n${table(["Cells", ...WRITE_COLS], rows)}\n`;
	}
	return md;
}

// The cross-format read section: one sub-table per container, columns = the libraries that read it
// (xlsb/ods have no ExcelJS column — it can't read them). Same time·memory cells as the main matrix.
function formatSection(results) {
	const sizeKeys = [...new Set(results.map((r) => r.sizeKey))];
	let md = "";
	for (const format of FORMATS) {
		const labels = FORMAT_READERS[format].map((r) => r.label);
		const rows = sizeKeys.map((sizeKey) => [
			`\`${sizeKey}\``,
			...labels.map((label) =>
				cellTimeMem(
					results.find(
						(r) => r.format === format && r.sizeKey === sizeKey && r.label === label,
					),
				),
			),
		]);
		md += `\n#### ${format}\n\n${table(["Cells", ...labels], rows)}\n`;
	}
	return md;
}

// The library size matrix — a clean production install of each library, from ../.cache/sizes.json
// (populated out-of-band by sizes.mjs, since it does real npm installs). Speed is one axis; the cost
// of the dependency footprint is the other, and the one openjsxl's zero-dep design targets.
function sizesSection(data) {
	if (!data) {
		return [
			"\nNo matching size measurements. Legacy or differently sourced results are omitted. Populate with:",
			"",
			"```sh",
			"node packages/bench/src/sizes.mjs   # writes packages/bench/.cache/sizes.json",
			"pnpm bench --render-only            # re-render with the size table filled in",
			"```",
			"",
		].join("\n");
	}
	// Sub-1-MB installs show 2 decimals so a ~0.45 MB library doesn't read as a lossy "0.4 MB".
	const fmtInstalled = (kb) => `${(kb / 1024).toFixed(kb < 1024 ? 2 : 1)} MB`;
	const rows = data.libs.map((l) => [
		`**${l.name}**`,
		`\`${l.version}\``,
		String(l.totalDeps),
		fmtBytes(l.ownBytes),
		fmtInstalled(l.installKB),
	]);
	const notes = data.libs
		.filter((l) => l.note)
		.map((l) => `\n_${l.name}: ${l.note}_\n`)
		.join("");
	return `
A clean production install (\`npm install --omit=dev --ignore-scripts\`) of each exact artifact — measured ${data.date} on
Node ${data.node.replace(/^v/, "")}. **Deps** counts the runtime packages the library pulls in
(transitive, excluding the library itself); **Own size** is the package's own unpacked tarball;
**Installed** is the library plus every runtime dependency on disk.
Comparator tarball bytes and npm metadata must match the artifact integrity in the recorded lockfile.

${table(["Library", "Version", "Deps", "Own size", "Installed"], rows)}
${notes}
openjsxl ships **zero third-party runtime dependencies** (both its packages are pure TypeScript over
platform Web APIs). The dependency counts above describe the measured distributions; SheetJS
0.20.3 also declares no runtime dependencies. *Installed* includes the library itself.
`;
}

function pythonSection(data) {
	if (!data) {
		return [
			"\nNo matching Python reference measurements. Legacy or different-fixture results are omitted.",
			"Reference numbers for Python readers/writers are gathered out-of-band (they need a",
			"Python environment, not npm). Run the companion script to populate this section:",
			"",
			"```sh",
			"python3 packages/bench/py/bench_py.py   # writes packages/bench/.cache/python.json",
			"pnpm bench --render-only                 # preserve the JS measurement date and versions",
			"```",
			"",
		].join("\n");
	}
	const cols = ["openpyxl", "python-calamine"];
	const versions = cols
		.map((name) => `${name} ${data.libraries?.[name] ?? "(version not recorded)"}`)
		.join("; ");
	let md = `\nMeasured out-of-band on: ${data.machine ?? "unknown"} · ${data.date ?? ""}\n
Libraries: ${versions}. Python CSV uses the recorded Python runtime's standard library.

> Reference only — **not** directly comparable to the JS numbers above. The Python runtime has a
> much lower baseline RSS than Node, openpyxl runs in its streaming \`read_only\`/\`write_only\` modes
> (tiny memory, unlike the in-memory JS APIs benchmarked above), and \`python-calamine\` is a native
> Rust binding. These are cross-language reference measurements, not same-runtime peers.
`;
	for (const op of ["read", "write"]) {
		const workloads = [
			...new Set(data.results.filter((r) => r.op === op).map((r) => r.workload)),
		];
		if (workloads.length === 0) continue;
		md += `\n**${op === "read" ? "Read" : "Write"}**\n`;
		for (const workload of workloads) {
			const sizes = [
				...new Set(
					data.results
						.filter((r) => r.op === op && r.workload === workload)
						.map((r) => r.sizeKey),
				),
			];
			const rows = sizes.map((sizeKey) => [
				`\`${sizeKey}\``,
				...cols.map((lib) => {
					const r = data.results.find(
						(x) =>
							x.op === op &&
							x.workload === workload &&
							x.sizeKey === sizeKey &&
							x.lib === lib,
					);
					return r ? `${fmtMs(r.timeMs)} · ${fmtBytes(r.peakRssMB * 1024 ** 2)}` : "—";
				}),
			]);
			md += `\n_${workload}_\n\n${table(["Cells", ...cols], rows)}\n`;
		}
	}
	// The cross-format read reference: calamine for xlsx/xlsb/ods, stdlib csv for csv.
	const formatRes = data.results.filter((r) => r.op === "read-format");
	if (formatRes.length) {
		const formats = [...new Set(formatRes.map((r) => r.format))];
		const sizeKeys = [...new Set(formatRes.map((r) => r.sizeKey))];
		md +=
			"\n**Read by format** — successful measurements on the matching fixtures: calamine (container formats) · stdlib `csv` (csv).\n\n";
		const rows = sizeKeys.map((sk) => [
			`\`${sk}\``,
			...formats.map((f) => {
				const r = formatRes.find((x) => x.format === f && x.sizeKey === sk);
				return r ? `${fmtMs(r.timeMs)} · ${fmtBytes(r.peakRssMB * 1024 ** 2)}` : "—";
			}),
		]);
		md += `${table(["Cells", ...formats], rows)}\n`;
	}
	return md;
}

/** Render recorded measurements without changing their date, machine, or library identity. */
export function renderReport(recordedResults, meta, supplemental = {}) {
	const results = validateReadResults(recordedResults);
	const { sizes, workloads, ops } = meta;
	const provenance = requireProvenance(meta.provenance);
	const date = provenance.measuredAt.slice(0, 10);
	const libraries = provenance.libraries
		.map((lib) => {
			const source = lib.source.startsWith("https://")
				? `[official distribution](${lib.source})`
				: `\`${lib.source}\``;
			const overrides = Object.entries(lib.overrides ?? {})
				.map(([name, version]) => `${name} ${version}`)
				.join(", ");
			return `**${lib.name}** \`${lib.version}\` (${source}${overrides ? `; override ${overrides}` : ""})`;
		})
		.join("; ");
	const sizesData = matchingSizeSnapshot(supplemental.sizes, provenance)
		? supplemental.sizes
		: undefined;
	const pythonData = matchingPythonSnapshot(supplemental.python, meta)
		? supplemental.python
		: undefined;
	// The cross-format read results are tagged separately so they don't leak into the main xlsx matrix
	// (which shares op/workload/label with the format phase's xlsx lane).
	const matrix = results.filter((r) => r.section !== "format-read");
	const formatReads = results.filter((r) => r.section === "format-read");
	const anyStyledWriteGap = matrix.some(
		(r) => r.op === "write" && r.workload === "styled" && r.label === "SheetJS" && !r.ok,
	);

	const md = `# openjsxl benchmarks

_Measured by \`pnpm bench\` on ${date}. Re-rendering preserves this measurement date._

**Machine:** ${machineLine(provenance.machine)}

Each number is the **median wall-time** (over warmed iterations) and the **peak process RSS of a
single cold call** (library load + one operation), measured in a **fresh, isolated Node process per
cell** so no library colours another's timing or memory. Workloads are 10 columns wide;
\`10k\`/\`100k\`/\`1M\` count cells. Read fixtures are authored once by ExcelJS (realistic shared
strings + a styles table) and read by all three JS readers. Full methodology and the fairness
caveats are at the bottom.

Libraries: ${libraries}. Peak RSS is absolute and includes loading the library.

**Source snapshot:** lockfile SHA-256 \`${provenance.lockfileSha256}\`; core bundle SHA-256
\`${provenance.bundles.core}\`. The raw result cache also records locked comparator artifact
integrities and fixture hashes.

## Library size
${sizesSection(sizesData)}
${ops.includes("read") ? `\n## Read\n\nParse a real \`.xlsx\` and materialize every cell value.\n${section(matrix, "read", sizes, workloads)}` : ""}
${ops.includes("write") ? `\n## Write\n\nSerialize N cells to \`.xlsx\` bytes. openjsxl's *streamed* column is fed a lazy row source (a\ngenerator) — the honest streaming case, where the full row array never exists in memory.\n${section(matrix, "write", sizes, workloads)}\n### Output file size\n${sizeSection(matrix, sizes, workloads)}` : ""}
${formatReads.length ? `\n## Read by format\n\nThe same \`numbers\` data authored in four container formats, each read by the libraries that support\nit and materialized cell-by-cell. xlsx reuses the ExcelJS-authored fixture; \`.xlsb\`/\`.ods\` are\nauthored by **SheetJS** (a real producer that writes those containers — the same author-then-read\nshape ExcelJS→xlsx already uses); \`.csv\` is written directly (plain text, no producer). ExcelJS reads\nonly xlsx/csv. Python reference rows, when available below, use the exact same fixture hashes.\nFormat compatibility is measured with the recorded producer version; old producer results are not reused.\n${formatSection(formatReads)}` : ""}
${failureSection(results)}

## Reference: Python
${pythonSection(pythonData)}

## Methodology

- **Isolation.** Every \`(library, operation, workload, size)\` cell runs in its own
  \`node --expose-gc\` child process, so a cell's memory and timing reflect that library alone.
- **Timing.** \`performance.now()\` around the operation; the reported value is the **median** of the
  measured iterations (2–5 depending on size), each preceded by a full GC, after warmup runs that
  are discarded.
- **Memory.** Peak \`process.memoryUsage().rss\` of the **first, cold run only** (which also serves as
  JIT warmup) — sampled on a 3 ms timer that fires during async awaits, plus a post-op read for
  synchronous ops. Measuring across warmed iterations would over-report: V8 doesn't return freed heap
  to the OS, so later iterations carry earlier ones' residue. The cold figure is what a single call
  costs; it is absolute, so it includes loading the library.
- **Equal work.** Every writer is handed the identical dataset; every reader parses the identical
  ExcelJS-authored file and materializes every cell into JS values. Outside worker timing, the
  harness compares each returned read checksum with the deterministic workload's expected sum,
  requiring a finite value within max(1e-9, 32 × Number.EPSILON × |expected|). Mismatches are excluded
  and listed above; this checksum is a sanity check, not proof of complete semantic equivalence.
  SheetJS writes with \`compression: true\` so its output is deflated like the
  others'.${anyStyledWriteGap ? "\n- **SheetJS styled write** is shown as —: emitting cell styles is a SheetJS Pro feature, absent from this community build, so measuring it would compare unequal work." : ""}
- **Not measured here.** CI-run benchmarks (too noisy) and internal micro-benchmarks. Python numbers
  are gathered out-of-band with a companion script and clearly marked as reference.

Regenerate with \`pnpm bench\` (add \`--quick\` for a fast smoke run). See
[\`packages/bench\`](../packages/bench) for the harness.
`;

	return md;
}

function readSupplement(path) {
	try {
		return existsSync(path) ? JSON.parse(readFileSync(path, "utf8")) : undefined;
	} catch {
		return undefined;
	}
}

/** Write docs/benchmarks.md from the collected results; returns the path written. */
export function writeReport(results, meta) {
	const md = renderReport(results, meta, {
		sizes: readSupplement(SIZES_CACHE),
		python: readSupplement(PYTHON_CACHE),
	});
	mkdirSync(fileURLToPath(new URL("../../../docs/", import.meta.url)), { recursive: true });
	writeFileSync(DOCS, md);
	return DOCS;
}
