import fc from "fast-check";
import { describe, it } from "vitest";
import {
	deterministicBytes,
	resolveOrTypedError,
	scalarRoundTrip,
	validRoundTrips,
} from "../writer-properties";

// F9.4 Half A — CI smoke. Fixed seed + modest numRuns so the writer property fuzzer runs
// deterministically inside the root gate in well under a second. The long local run (packages/fuzz
// `fuzz` script) replays the SAME properties with a far larger numRuns; see long-run.test.ts.

const SMOKE = { seed: 0x0f9_4a, numRuns: 60, endOnFailure: true } as const;

describe("writer fuzz (Half A) — smoke", () => {
	it("P1 resolve-or-typed-error: writeXlsx round-trips OR throws XlsxError('invalid-input')", async () => {
		await fc.assert(resolveOrTypedError, SMOKE);
	});

	it("P4 valid input round-trips: a valid workbook always writes AND re-opens", async () => {
		await fc.assert(validRoundTrips, SMOKE);
	});

	it("P2 deterministic bytes: identical input → identical output", async () => {
		await fc.assert(deterministicBytes, SMOKE);
	});

	it("P3 scalar round-trip: every written scalar reads back unchanged", async () => {
		await fc.assert(scalarRoundTrip, SMOKE);
	});
});
