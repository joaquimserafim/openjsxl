// Half B — the mutation campaign: enumerate the fixture corpus, mutate each seed under a deterministic
// PRNG, replay every mutant against all readers, and collect crashers. Two modes:
//   • fixed-count (`mutantsPerFixture`) — deterministic, bounded; used by the CI smoke suite;
//   • budgeted (`budgetMs` + `rssLimitBytes`) — the long local run, which stops on wall-clock or an RSS
//     ceiling (so a leak or a runaway allocation aborts the run instead of the machine).
// A crasher keeps its exact mutated bytes and the integer seed that produced it, so it is replayable
// and promotable to a committed regression fixture.

import { memoryUsage } from "node:process";
import { listFixtures, loadFixture } from "@openjsxl/fixtures";
import { mutate } from "./mutate";
import { Prng } from "./prng";
import { type OpenerOutcome, replayAll } from "./replay";

export interface Crasher {
	readonly fixture: string;
	readonly seed: number;
	readonly crashes: readonly OpenerOutcome[];
	readonly bytes: Uint8Array;
}

export interface CampaignReport {
	readonly mutants: number;
	readonly crashers: readonly Crasher[];
	readonly fixtures: readonly string[];
	readonly stoppedBy: "count" | "budget" | "rss";
	/** Verdict tally across every opener-outcome (mutants × openers). Diagnostic (see `changed`). */
	readonly tally: { readonly ok: number; readonly typed: number; readonly crash: number };
	/**
	 * How many mutants' bytes actually DIFFER from their seed fixture. This — not `tally.typed` — is
	 * the real "the engine bites" signal: the cross-format openers reject other-format fixtures typed
	 * even with zero mutation, so `tally.typed > 0` is vacuous, whereas `changed` is 0 iff `mutate()`
	 * is a no-op. The smoke asserts `changed === mutants`.
	 */
	readonly changed: number;
}

export interface CampaignOptions {
	/** Fixed-count mode: exactly this many mutants per fixture (deterministic). Ignored if `budgetMs` set. */
	readonly mutantsPerFixture?: number;
	/** Budgeted mode: keep mutating until this many wall-clock ms elapse. */
	readonly budgetMs?: number;
	/** Budgeted mode: abort if RSS exceeds this many bytes (guards against a leak/OOM). */
	readonly rssLimitBytes?: number;
	/** Base seed folded into every mutant's PRNG — same base ⇒ identical campaign. */
	readonly baseSeed?: number;
	/** Restrict to these fixture names; defaults to the whole committed corpus. */
	readonly fixtures?: readonly string[];
	/** Progress callback (long run) — invoked every `progressEvery` mutants. */
	readonly onProgress?: (done: number, crashers: number) => void;
	readonly progressEvery?: number;
}

// A deterministic per-mutant seed from the fixture index and iteration — spreads well and is stable.
function mutantSeed(base: number, fixtureIndex: number, k: number): number {
	return (base ^ (fixtureIndex * 0x1eb3 + k * 0x9e37)) | 0;
}

async function loadCorpus(
	names: readonly string[],
): Promise<{ name: string; bytes: Uint8Array }[]> {
	const out: { name: string; bytes: Uint8Array }[] = [];
	for (const name of names) out.push({ name, bytes: await loadFixture(name) });
	return out;
}

/**
 * Run a mutation campaign and return every crasher found. Deterministic in fixed-count mode; in
 * budgeted mode the SET of mutants depends on wall-clock but each individual mutant is still
 * reproducible from `{fixture, seed}`.
 */
export async function runCampaign(options: CampaignOptions = {}): Promise<CampaignReport> {
	const base = options.baseSeed ?? 0x0f94;
	const names = options.fixtures ?? listFixtures();
	const corpus = await loadCorpus(names);
	const crashers: Crasher[] = [];
	let mutants = 0;
	let changed = 0;
	const tally = { ok: 0, typed: 0, crash: 0 };

	// One mutate + replay + bookkeeping step: mutate the seed, count whether it actually changed,
	// replay against every reader, and tally verdicts / collect crashers.
	const step = async (fixture: string, seedBytes: Uint8Array, seed: number): Promise<void> => {
		const mutant = mutate(seedBytes, new Prng(seed));
		if (!bytesEqual(mutant, seedBytes)) changed++;
		const { outcomes } = await replayAll(mutant);
		for (const o of outcomes) tally[o.verdict]++;
		const crashes = outcomes.filter((o) => o.verdict === "crash");
		if (crashes.length > 0) crashers.push({ fixture, seed, crashes, bytes: mutant });
		mutants++;
	};

	if (options.budgetMs !== undefined) {
		// Budgeted mode: round-robin the corpus until the clock (or RSS) says stop.
		const deadline = performance.now() + options.budgetMs;
		const rssLimit = options.rssLimitBytes;
		const every = options.progressEvery ?? 500;
		let k = 0;
		let stoppedBy: "budget" | "rss" = "budget";
		while (performance.now() < deadline) {
			for (let f = 0; f < corpus.length; f++) {
				const entry = corpus[f];
				if (entry === undefined) continue;
				await step(entry.name, entry.bytes, mutantSeed(base, f, k));
				if (mutants % every === 0) options.onProgress?.(mutants, crashers.length);
				if (rssLimit !== undefined && memoryUsage().rss > rssLimit) {
					stoppedBy = "rss";
					return { mutants, crashers, fixtures: names, stoppedBy, tally, changed };
				}
			}
			k++;
			if (performance.now() >= deadline) break;
		}
		return { mutants, crashers, fixtures: names, stoppedBy, tally, changed };
	}

	// Fixed-count mode: exactly `per` mutants per fixture — fully deterministic.
	const per = options.mutantsPerFixture ?? 20;
	for (let f = 0; f < corpus.length; f++) {
		const entry = corpus[f];
		if (entry === undefined) continue;
		for (let k = 0; k < per; k++) await step(entry.name, entry.bytes, mutantSeed(base, f, k));
	}
	return { mutants, crashers, fixtures: names, stoppedBy: "count", tally, changed };
}

// Byte-equality of two buffers (used to detect a no-op mutation).
function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
	return true;
}
