// Benchmark orchestrator. Authors read fixtures once, then runs every (library, op, workload, size)
// cell in its OWN child process, strictly one at a time — running cells concurrently would let them
// steal each other's CPU and memory and make every number a lie. Collects the JSON each worker
// emits and hands the whole result set to the markdown reporter.
//
//   node src/run.mjs                       # full matrix → docs/benchmarks.md
//   node src/run.mjs --quick               # 10k only, 1 iter — a fast smoke test of the harness
//   node src/run.mjs --sizes 10k,100k --ops write --workloads numbers

import { spawn } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { ensureFormatFixture, FORMAT_READERS, FORMAT_WORKLOAD, FORMATS } from "./formats.mjs";
import { writeReport } from "./report.mjs";
import { COLS, SIZES, WORKLOADS } from "./workloads.mjs";

const HERE = new URL("./", import.meta.url);
const CACHE = fileURLToPath(new URL("../.cache/", HERE));
const WORKER = fileURLToPath(new URL("./worker.mjs", HERE));
// The raw JS results are cached so the report can be re-rendered without re-running the matrix —
// used by `--render-only` to merge the out-of-band Python numbers cheaply after they're gathered.
const JS_RESULTS = `${CACHE}js-results.json`;

// Per-size iteration policy: small workloads get more samples (cheap, less noise); 1M gets fewer
// (each run is seconds and the signal is already strong).
const ITER_POLICY = {
	"10k": { warmup: 2, iters: 5 },
	"100k": { warmup: 1, iters: 3 },
	"1M": { warmup: 1, iters: 2 },
};

// Read is compared across the three JS readers; write adds openjsxl's streamed path.
const READ_RUNNERS = [
	{ id: "openjsxl", label: "openjsxl", op: "read" },
	{ id: "exceljs", label: "ExcelJS", op: "read" },
	{ id: "xlsx", label: "SheetJS", op: "read" },
];
const WRITE_RUNNERS = [
	{ id: "openjsxl", label: "openjsxl (buffered)", op: "write" },
	{ id: "openjsxl", label: "openjsxl (streamed)", op: "write-stream" },
	{ id: "exceljs", label: "ExcelJS", op: "write" },
	{ id: "xlsx", label: "SheetJS", op: "write" },
];

function parseArgs(argv) {
	const opts = {
		sizes: null,
		workloads: null,
		ops: null,
		quick: false,
		iters: null,
		warmup: null,
		renderOnly: false,
	};
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "--quick") opts.quick = true;
		else if (a === "--render-only") opts.renderOnly = true;
		else if (a === "--sizes") opts.sizes = argv[++i].split(",");
		else if (a === "--workloads") opts.workloads = argv[++i].split(",");
		else if (a === "--ops") opts.ops = argv[++i].split(",");
		else if (a === "--iters") opts.iters = Number(argv[++i]);
		else if (a === "--warmup") opts.warmup = Number(argv[++i]);
	}
	return opts;
}

// Spawn one worker cell in a fresh process and resolve its parsed JSON result.
function runCell(cfg) {
	return new Promise((resolve) => {
		const child = spawn(
			process.execPath,
			["--expose-gc", "--max-old-space-size=8192", WORKER, JSON.stringify(cfg)],
			{ stdio: ["ignore", "pipe", "inherit"] },
		);
		let out = "";
		child.stdout.on("data", (chunk) => {
			out += chunk;
		});
		child.on("close", () => {
			const line = out.trim().split("\n").filter(Boolean).pop();
			if (!line) return resolve({ ok: false, ...cfg, reason: "no output (crash / OOM)" });
			try {
				resolve(JSON.parse(line));
			} catch {
				resolve({ ok: false, ...cfg, reason: "unparseable output" });
			}
		});
		child.on("error", (err) => resolve({ ok: false, ...cfg, reason: err.message }));
	});
}

// Author the canonical read fixture for (workload, size) with ExcelJS — a neutral, realistic
// producer (shared strings, a styles table), so every JS reader parses the exact same real file.
// Cached; regenerated only when absent.
async function ensureFixture(workload, size) {
	const path = `${CACHE}read-${workload}-${size.key}.xlsx`;
	if (existsSync(path)) return path;
	const { buildDataset } = await import("./workloads.mjs");
	const exceljs = await import("./adapters/exceljs.mjs");
	process.stdout.write(`  · authoring fixture ${workload}/${size.key} (ExcelJS) …\n`);
	const bytes = await exceljs.write(buildDataset(workload, size.rows), workload);
	writeFileSync(path, bytes);
	return path;
}

async function main() {
	const opts = parseArgs(process.argv.slice(2));
	mkdirSync(CACHE, { recursive: true });

	// Re-render the report from the last run's cached results (+ any Python numbers) without paying
	// for the whole matrix again.
	if (opts.renderOnly) {
		if (!existsSync(JS_RESULTS)) throw new Error("no cached results — run `pnpm bench` first");
		const cached = JSON.parse(readFileSync(JS_RESULTS, "utf8"));
		const outPath = writeReport(cached.results, cached.meta);
		process.stdout.write(`Re-rendered ${outPath} from cached results\n`);
		return;
	}

	const sizeKeys = opts.quick ? ["10k"] : (opts.sizes ?? SIZES.map((s) => s.key));
	const sizes = SIZES.filter((s) => sizeKeys.includes(s.key));
	const workloads = opts.workloads ?? WORKLOADS;
	const ops = opts.ops ?? ["read", "write"];

	const results = [];
	for (const size of sizes) {
		const policy = ITER_POLICY[size.key] ?? { warmup: 1, iters: 3 };
		const iters = opts.quick ? 1 : (opts.iters ?? policy.iters);
		const warmup = opts.quick ? 0 : (opts.warmup ?? policy.warmup);

		for (const workload of workloads) {
			// READ cells (each needs the shared fixture authored first).
			if (ops.includes("read")) {
				const fixture = await ensureFixture(workload, size);
				for (const r of READ_RUNNERS) {
					const cfg = {
						lib: r.id,
						op: r.op,
						workload,
						rows: size.rows,
						iters,
						warmup,
						fixture,
					};
					process.stdout.write(
						`▶ read  · ${workload.padEnd(7)} · ${size.key.padEnd(4)} · ${r.label} … `,
					);
					const res = await runCell(cfg);
					res.label = r.label;
					res.sizeKey = size.key;
					results.push(res);
					report(res);
				}
			}
			// WRITE cells.
			if (ops.includes("write")) {
				for (const r of WRITE_RUNNERS) {
					const cfg = { lib: r.id, op: r.op, workload, rows: size.rows, iters, warmup };
					process.stdout.write(
						`▶ write · ${workload.padEnd(7)} · ${size.key.padEnd(4)} · ${r.label} … `,
					);
					const res = await runCell(cfg);
					res.label = r.label;
					res.sizeKey = size.key;
					results.push(res);
					report(res);
				}
			}
		}
	}

	// Cross-format READ phase: the same `numbers` data in xlsx/xlsb/ods/csv, read by each library
	// that supports the format. Runs unless read is explicitly excluded. Tagged section:"format-read"
	// so the reporter keeps it separate from the main xlsx read matrix.
	if (ops.includes("read")) {
		for (const size of sizes) {
			const policy = ITER_POLICY[size.key] ?? { warmup: 1, iters: 3 };
			const iters = opts.quick ? 1 : (opts.iters ?? policy.iters);
			const warmup = opts.quick ? 0 : (opts.warmup ?? policy.warmup);
			const xlsxFixture = await ensureFixture(FORMAT_WORKLOAD, size);
			for (const format of FORMATS) {
				const fixture = await ensureFormatFixture(CACHE, size, format, xlsxFixture);
				for (const r of FORMAT_READERS[format]) {
					const cfg = {
						lib: r.id,
						op: "read",
						workload: FORMAT_WORKLOAD,
						rows: size.rows,
						iters,
						warmup,
						fixture,
						format,
					};
					process.stdout.write(
						`▶ read· ${format.padEnd(4)} · ${size.key.padEnd(4)} · ${r.label} … `,
					);
					const res = await runCell(cfg);
					res.label = r.label;
					res.sizeKey = size.key;
					res.format = format;
					res.section = "format-read";
					results.push(res);
					report(res);
				}
			}
		}
	}

	const meta = { cols: COLS, sizes, workloads, ops };
	writeFileSync(JS_RESULTS, JSON.stringify({ meta, results }, null, 2));
	const outPath = writeReport(results, meta);
	process.stdout.write(`\nWrote ${outPath}\n`);
}

function report(res) {
	if (!res.ok) {
		process.stdout.write(`n/a (${res.reason})\n`);
		return;
	}
	const mb = (res.peakRss / 1024 ** 2).toFixed(0);
	const t =
		res.timeMs < 1000 ? `${res.timeMs.toFixed(0)} ms` : `${(res.timeMs / 1000).toFixed(2)} s`;
	process.stdout.write(`${t} · peak ${mb} MB\n`);
}

main().catch((err) => {
	process.stderr.write(`${err?.stack ?? err}\n`);
	process.exit(1);
});
