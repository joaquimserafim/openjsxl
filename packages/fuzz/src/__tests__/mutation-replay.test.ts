import { describe, expect, it } from "vitest";
import { runCampaign } from "../campaign";

// F9.4 Half B — CI smoke. A fixed-count, fixed-seed mutation campaign over the whole committed corpus.
// The invariant: no mutant makes any reader CRASH (throw a non-XlsxError). Deterministic and bounded,
// so it runs in the root gate; the long local run (`fuzz` script) is budgeted instead. No wall-clock
// assert here — CI machines vary; only the crash count is asserted.

describe("mutation replay (Half B) — smoke", () => {
	it("no mutant crashes any reader (resolve OR typed XlsxError only)", async () => {
		const report = await runCampaign({ mutantsPerFixture: 12, baseSeed: 0x0f94 });
		if (report.crashers.length > 0) {
			// Surface enough to reproduce: fixture + integer seed + which opener + the error.
			const detail = report.crashers
				.slice(0, 10)
				.map(
					(c) =>
						`  ${c.fixture} seed=${c.seed}: ${c.crashes.map((x) => `${x.opener} → ${x.error}`).join("; ")}`,
				)
				.join("\n");
			throw new Error(
				`${report.crashers.length}/${report.mutants} mutants crashed a reader:\n${detail}`,
			);
		}
		expect(report.crashers).toEqual([]);
		// Sanity: the campaign actually ran over the corpus.
		expect(report.mutants).toBeGreaterThan(0);
		expect(report.fixtures.length).toBeGreaterThan(10);
		// The mutations must actually BITE. `tally.typed` is NOT a valid signal (the cross-format openers
		// reject other-format fixtures typed even with zero mutation) — the sound guard is that a mutant's
		// bytes DIFFER from its seed. A no-op `mutate()` gives `changed === 0` and fails here; the >90%
		// threshold (not `=== mutants`) tolerates the rare coincidental no-op (e.g. a byteSet writing an
		// identical byte) without weakening the no-op-engine guard.
		expect(report.changed).toBeGreaterThan(report.mutants * 0.9);
		expect(report.tally.crash).toBe(0);
	});
});
