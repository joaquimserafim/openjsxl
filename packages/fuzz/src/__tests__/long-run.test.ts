import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { runCampaign } from "../campaign";
import { writeCrasher } from "../triage";
import {
	deterministicBytes,
	resolveOrTypedError,
	scalarRoundTrip,
	validRoundTrips,
} from "../writer-properties";

// F9.4 — the LONG local run. Gated behind FUZZ_LONG=1 so the root gate never runs it (the smoke suites
// cover CI). Invoke via `pnpm --filter @openjsxl/fuzz fuzz`, tuned by env:
//   FUZZ_MS      wall-clock budget for the mutation campaign (default 60000)
//   FUZZ_RUNS    fast-check numRuns for each writer property (default 5000)
//   FUZZ_RSS_MB  RSS ceiling for the campaign, MB (default 2048)
//   FUZZ_SEED    fast-check seed (default 42; set for a fresh exploration)
// A mutation crasher is minimized and its reproducer written to the gitignored crashers/ dir before
// the run fails; a writer-property failure is shrunk and reported by fast-check.

const LONG = process.env.FUZZ_LONG === "1" || process.env.FUZZ_LONG === "true";
const num = (k: string, d: number): number => {
	const v = Number(process.env[k]);
	return Number.isFinite(v) && v > 0 ? v : d;
};

const BUDGET_MS = num("FUZZ_MS", 60_000);
const RUNS = num("FUZZ_RUNS", 5_000);
const RSS_BYTES = num("FUZZ_RSS_MB", 2_048) * 1024 * 1024;
const SEED = num("FUZZ_SEED", 42);
const SLACK_MS = 120_000;

describe.skipIf(!LONG)("F9.4 long run", () => {
	it(
		"Half A — writer properties at scale",
		async () => {
			const cfg = { seed: SEED, numRuns: RUNS, endOnFailure: true } as const;
			await fc.assert(resolveOrTypedError, cfg);
			await fc.assert(validRoundTrips, cfg);
			await fc.assert(deterministicBytes, cfg);
			await fc.assert(scalarRoundTrip, cfg);
		},
		RUNS * 40 + SLACK_MS,
	);

	it(
		"Half B — mutation campaign under a wall-clock + RSS budget",
		async () => {
			const report = await runCampaign({
				budgetMs: BUDGET_MS,
				rssLimitBytes: RSS_BYTES,
				baseSeed: SEED,
				progressEvery: 1_000,
				onProgress: (done, crashers) => {
					console.log(
						`[fuzz] ${done} mutants, ${crashers} crashers, rss=${Math.round(process.memoryUsage().rss / 1e6)}MB`,
					);
				},
			});
			console.log(
				`[fuzz] done: ${report.mutants} mutants (${report.changed} changed) over ${report.fixtures.length} fixtures, stoppedBy=${report.stoppedBy}, ${report.crashers.length} crashers; verdicts ok=${report.tally.ok} typed=${report.tally.typed} crash=${report.tally.crash}`,
			);
			if (report.crashers.length > 0) {
				const paths: string[] = [];
				for (const c of report.crashers.slice(0, 50)) paths.push(await writeCrasher(c));
				throw new Error(
					`${report.crashers.length} crashers; reproducers written:\n${paths.join("\n")}`,
				);
			}
			// Stopping on the RSS ceiling is a GRACEFUL early stop, not a failure: each mutant's Workbook
			// is transient (GC-eligible once replayAll returns) and nothing accumulates across mutants, so
			// the high-water mark is V8 heap slack under a tight allocation loop, not a reader leak (a true
			// leak grows unbounded regardless of budget). Only crashers fail the run; RSS just warns.
			if (report.stoppedBy === "rss") {
				console.warn(
					`[fuzz] stopped early on the RSS ceiling after ${report.mutants} mutants (heap slack under the loop; raise FUZZ_RSS_MB for a longer run)`,
				);
			}
			expect(report.mutants).toBeGreaterThan(0);
		},
		BUDGET_MS + SLACK_MS,
	);
});
