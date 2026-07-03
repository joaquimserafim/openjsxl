// Memory harness for the streaming writer (F5.1). Streams sheets of increasing size through
// streamXlsx, discarding the output, and reports peak resident-set growth. If streaming is truly
// constant-memory, peak RSS should be roughly FLAT across row counts — a linear climb would mean the
// sheet is being buffered somewhere. Run from the repo root after `pnpm build`:
//
//     node --expose-gc scripts/stream-memory.mjs
//
// (--expose-gc is optional; it just makes the baseline cleaner.)

import { streamXlsx } from "../packages/core/dist/index.js";

async function* rows(n) {
	// A DB-cursor-shaped async source: never materializes the whole sheet.
	for (let i = 1; i <= n; i++) yield [i, `label ${i}`, i * 1.5, i % 2 === 0];
}

async function run(n) {
	const reader = streamXlsx({ sheets: [{ name: "S", rows: rows(n) }] }).getReader();
	let bytes = 0;
	let chunks = 0;
	let liveHeapPeak = 0; // max LIVE heap (post-GC) sampled mid-stream — the true working set
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		bytes += value.length;
		// Sample the retained set periodically: force GC then read heapUsed (garbage excluded).
		if (globalThis.gc && ++chunks % 200 === 0) {
			globalThis.gc();
			const live = process.memoryUsage().heapUsed;
			if (live > liveHeapPeak) liveHeapPeak = live;
		}
	}
	return { n, mb: bytes / 1e6, liveMb: liveHeapPeak / 1e6 };
}

console.log("rows        output    peak LIVE heap (post-GC)");
for (const n of [10_000, 100_000, 500_000, 1_000_000]) {
	const r = await run(n);
	console.log(
		`${r.n.toLocaleString().padStart(9)}  ${r.mb.toFixed(1).padStart(6)} MB  ${r.liveMb.toFixed(1).padStart(7)} MB`,
	);
}
console.log(
	"\nPeak LIVE heap should stay roughly flat as rows grow (independent of file size) — that is\nconstant-memory streaming. Run with --expose-gc for the post-GC sampling.",
);
