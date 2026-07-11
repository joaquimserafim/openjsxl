// One benchmark cell, measured in a fresh, isolated process (spawned by run.mjs). Isolation is the
// whole point: each (library, op, workload, size) gets its own clean V8 heap, so one library's
// allocations never colour another's memory or timing numbers. Started with --expose-gc so the heap
// can be collected between iterations, and the per-op peak reflects that op alone.
//
// Config arrives as a single JSON argument. Result is a single JSON line on stdout.

import { readFileSync } from "node:fs";
import { median } from "./stats.mjs";
import { buildDataset, datasetGenerator } from "./workloads.mjs";

const cfg = JSON.parse(process.argv[2]);
const { lib, op, workload, rows, iters, warmup, fixture, format = "xlsx" } = cfg;

function rss() {
	return process.memoryUsage().rss;
}

// Force a full GC if exposed; used to settle the heap before each measured iteration so a prior
// iteration's garbage can't inflate this one's peak.
function collect() {
	if (typeof global.gc === "function") {
		global.gc();
		global.gc();
	}
}

async function main() {
	const adapter = await import(`./adapters/${lib}.mjs`);

	// Prepare inputs OUTSIDE the timed region, so dataset generation / fixture I/O never counts.
	let bytes;
	let dataset;
	if (op === "read") bytes = readFileSync(fixture);
	else if (op === "write") dataset = buildDataset(workload, rows);
	// write-stream builds a fresh lazy generator per iteration (no materialized array by design).

	const runOnce = async () => {
		if (op === "read") return await adapter.read(bytes, format);
		if (op === "write") return await adapter.write(dataset, workload);
		return await adapter.writeStream(datasetGenerator(workload, rows));
	};

	// Peak RSS is measured on the FIRST, cold run only — a single call's honest working set (module
	// load + one operation). Measuring it across warmed iterations would inflate it: V8 doesn't return
	// freed heap to the OS, so each later iteration carries the residue of the ones before, reporting
	// more memory than any single call actually costs. The cold run doubles as JIT warmup #1; a coarse
	// sampler (fires during async awaits) plus a post-op read (for synchronous ops) catch the peak.
	collect();
	let peakRss = rss();
	const sampler = setInterval(() => {
		const r = rss();
		if (r > peakRss) peakRss = r;
	}, 3);
	sampler.unref();
	let last = await runOnce();
	clearInterval(sampler);
	{
		const post = rss();
		if (post > peakRss) peakRss = post;
	}

	// Any further warmup runs (JIT), then the timed iterations — wall-time only; memory stays the cold
	// figure above. The cold run already served as one warmup, so only warmup-1 extra are needed.
	for (let i = 1; i < warmup; i++) last = await runOnce();

	const times = [];
	for (let i = 0; i < iters; i++) {
		collect();
		const t0 = performance.now();
		last = await runOnce();
		const t1 = performance.now();
		times.push(t1 - t0);
	}

	// Size/shape of the produced or consumed data, for the report + a sanity check on equal work.
	// Reads report a checksum sink; writes report the produced byte length (Uint8Array or {length}).
	let outBytes;
	let sink;
	if (op === "read") sink = last;
	else outBytes = last.length;

	process.stdout.write(
		`${JSON.stringify({
			ok: true,
			lib,
			op,
			workload,
			rows,
			timeMs: median(times),
			times,
			peakRss,
			outBytes,
			sink,
		})}\n`,
	);
}

main().catch((err) => {
	const message = err instanceof Error ? err.message : String(err);
	process.stdout.write(
		`${JSON.stringify({ ok: false, lib, op, workload, rows, reason: message })}\n`,
	);
	process.exitCode = 0; // a failed cell is reported, not fatal to the run
});
