// Triage — turn a raw crasher into something committable: minimize the bytes to the smallest buffer
// that still crashes, then write a reproducer (bytes + seed + trace) into the gitignored `crashers/`
// dir. A confirmed, minimized crasher is then promoted BY HAND to `packages/fixtures/data/edge-*` with
// a verbatim-read regression (the data/README checklist) — never auto-committed.

import { mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type { Crasher } from "./campaign";
import { replayAll } from "./replay";

const crashersDir = fileURLToPath(new URL("../crashers/", import.meta.url));

/** Does this buffer still crash the given opener? (The minimization invariant.) */
async function stillCrashes(bytes: Uint8Array, opener: string): Promise<boolean> {
	const { outcomes } = await replayAll(bytes);
	return outcomes.some((o) => o.opener === opener && o.verdict === "crash");
}

/**
 * Chunk-removal minimization (a ddmin-lite): repeatedly try deleting contiguous chunks, keeping any
 * deletion that preserves the crash on `opener`, halving the chunk size when a full pass makes no
 * progress. Converges on a near-1-minimal buffer. Bounded by `maxPasses` so a pathological case can't
 * loop forever.
 */
export async function minimize(
	bytes: Uint8Array,
	opener: string,
	maxPasses = 24,
): Promise<Uint8Array> {
	let current = bytes;
	let chunk = Math.max(1, Math.floor(current.length / 2));
	for (let pass = 0; pass < maxPasses && chunk >= 1; pass++) {
		let progressed = false;
		for (let start = 0; start < current.length; start += chunk) {
			const candidate = new Uint8Array(
				current.length - Math.min(chunk, current.length - start),
			);
			candidate.set(current.subarray(0, start), 0);
			candidate.set(current.subarray(start + chunk), start);
			if (candidate.length > 0 && (await stillCrashes(candidate, opener))) {
				current = candidate;
				progressed = true;
			}
		}
		if (!progressed) chunk = Math.floor(chunk / 2);
	}
	return current;
}

/**
 * Write a crasher's reproducer into the gitignored `crashers/` dir: the raw bytes, a minimized copy,
 * and a `.txt` trace naming the seed fixture, integer seed, crashing opener(s), and error. Returns the
 * base path written. Node-only (long-run harness); never called from the CI smoke.
 */
export async function writeCrasher(crasher: Crasher): Promise<string> {
	mkdirSync(crashersDir, { recursive: true });
	const opener = crasher.crashes[0]?.opener ?? "unknown";
	const base = `${crasher.fixture}.seed${crasher.seed >>> 0}.${opener}`;
	const minimized = await minimize(crasher.bytes, opener);
	writeFileSync(`${crashersDir}${base}.bin`, crasher.bytes);
	writeFileSync(`${crashersDir}${base}.min.bin`, minimized);
	const trace = [
		`seed fixture : ${crasher.fixture}`,
		`mutant seed  : ${crasher.seed >>> 0}`,
		`raw bytes    : ${crasher.bytes.length}`,
		`min bytes    : ${minimized.length}`,
		"crashes:",
		...crasher.crashes.map((c) => `  ${c.opener} → ${c.error ?? ""}`),
	].join("\n");
	writeFileSync(`${crashersDir}${base}.txt`, trace);
	return `${crashersDir}${base}`;
}
