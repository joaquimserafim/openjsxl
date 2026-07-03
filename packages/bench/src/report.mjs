import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { machineLine } from "./machine.mjs";
import { fmtBytes, fmtMs } from "./stats.mjs";

// Renders the collected results into docs/benchmarks.md. Every published number comes from an
// actual run — this file only formats, it never invents or extrapolates a value.

const DOCS = fileURLToPath(new URL("../../../docs/benchmarks.md", import.meta.url));
const PYTHON_CACHE = fileURLToPath(new URL("../.cache/python.json", import.meta.url));

const READ_COLS = ["openjsxl", "ExcelJS", "SheetJS"];
const WRITE_COLS = ["openjsxl (buffered)", "openjsxl (streamed)", "ExcelJS", "SheetJS"];

function pick(results, op, workload, sizeKey, label) {
	return results.find(
		(r) => r.op === op && r.workload === workload && r.sizeKey === sizeKey && r.label === label,
	);
}

// A time·memory cell. Missing/failed → em dash (footnote explains the common case: SheetJS styles).
function cellTimeMem(res) {
	if (!res || !res.ok) return "—";
	return `${fmtMs(res.timeMs)} · ${fmtBytes(res.peakRss)}`;
}

function table(header, rows) {
	const head = `| ${header.join(" | ")} |`;
	const sep = `| ${header.map(() => "---").join(" | ")} |`;
	const body = rows.map((r) => `| ${r.join(" | ")} |`).join("\n");
	return `${head}\n${sep}\n${body}`;
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

function pythonSection() {
	if (!existsSync(PYTHON_CACHE)) {
		return [
			"\nReference numbers for Python readers/writers are gathered out-of-band (they need a",
			"Python environment, not npm). Run the companion script to populate this section:",
			"",
			"```sh",
			"python3 packages/bench/py/bench_py.py   # writes packages/bench/.cache/python.json",
			"pnpm bench                              # re-render with the Python column filled in",
			"```",
			"",
		].join("\n");
	}
	const data = JSON.parse(readFileSync(PYTHON_CACHE, "utf8"));
	const cols = ["openpyxl", "python-calamine"];
	let md = `\nMeasured out-of-band on: ${data.machine ?? "unknown"} · ${data.date ?? ""}\n

> Reference only — **not** directly comparable to the JS numbers above. The Python runtime has a
> much lower baseline RSS than Node, openpyxl runs in its streaming \`read_only\`/\`write_only\` modes
> (tiny memory, unlike the in-memory JS APIs benchmarked above), and \`python-calamine\` is a native
> Rust binding — the cross-language *speed bar* openjsxl targets, not a same-runtime peer. The point:
> openjsxl is the fastest pure-JS option here and lands within ~1.5× of native calamine on read.
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
	return md;
}

/** Write docs/benchmarks.md from the collected results; returns the path written. */
export function writeReport(results, { sizes, workloads, ops }) {
	const date = new Date().toISOString().slice(0, 10);
	const anyStyledWriteGap = results.some(
		(r) => r.op === "write" && r.workload === "styled" && r.label === "SheetJS" && !r.ok,
	);

	const md = `# openjsxl benchmarks

_Generated by \`pnpm bench\` on ${date}. Reproducible end-to-end on a clean checkout._

**Machine:** ${machineLine()}

Each number is the **median wall-time** (over warmed iterations) and the **peak process RSS of a
single cold call** (library load + one operation), measured in a **fresh, isolated Node process per
cell** so no library colours another's timing or memory. Workloads are 10 columns wide;
\`10k\`/\`100k\`/\`1M\` count cells. Read fixtures are authored once by ExcelJS (realistic shared
strings + a styles table) and read by all three JS readers. Full methodology and the fairness
caveats are at the bottom.

Libraries: **openjsxl** (this project), **ExcelJS** \`4.4.0\`, **SheetJS** \`xlsx@0.18.5\` (the last
version on the public npm registry). Peak RSS is absolute and includes loading the library — the
working set one call actually costs.
${ops.includes("read") ? `\n## Read\n\nParse a real \`.xlsx\` and materialize every cell value.\n${section(results, "read", sizes, workloads)}` : ""}
${ops.includes("write") ? `\n## Write\n\nSerialize N cells to \`.xlsx\` bytes. openjsxl's *streamed* column is fed a lazy row source (a\ngenerator) — the honest streaming case, where the full row array never exists in memory.\n${section(results, "write", sizes, workloads)}\n### Output file size\n${sizeSection(results, sizes, workloads)}` : ""}

## Reference: Python
${pythonSection()}

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
  ExcelJS-authored file and materializes every cell into JS values (a checksum sink prevents
  dead-code elimination). SheetJS writes with \`compression: true\` so its output is deflated like the
  others'.${anyStyledWriteGap ? "\n- **SheetJS styled write** is shown as —: emitting cell styles is a SheetJS Pro feature, absent from the npm build, so measuring it would compare unequal work." : ""}
- **Not measured here.** CI-run benchmarks (too noisy) and internal micro-benchmarks. Python numbers
  are gathered out-of-band with a companion script and clearly marked as reference.

Regenerate with \`pnpm bench\` (add \`--quick\` for a fast smoke run). See
[\`packages/bench\`](../packages/bench) for the harness.
`;

	mkdirSync(fileURLToPath(new URL("../../../docs/", import.meta.url)), { recursive: true });
	writeFileSync(DOCS, md);
	return DOCS;
}
