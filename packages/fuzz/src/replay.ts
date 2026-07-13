// Half B — replay a (mutated) byte buffer against every reader and classify the outcome. The single
// invariant every opener must uphold on ANY bytes: it either RESOLVES (returns a workbook / a detected
// format) or throws a typed {@link XlsxError}. A TypeError / RangeError / bare throw / rejection with
// any other error is a CRASH — a reader that trusted a corrupted length, offset, or count.
//
// Opening is NOT enough: `openXlsx` decompresses each sheet's XML up front but defers the A1 / shared-
// string / style SAX parse to the first cell access, so a mutant that yields decodable-but-malformed
// sheet XML would be scored `ok` if we only opened it. Every workbook opener therefore has its cells
// MATERIALIZED (drain every sheet's `rows()`) so the sheet parser is actually exercised. This is where
// the `xmlText` mutator's `count=/ref=/r=/span=` blow-ups land for the flagship xlsx reader.

import {
	detectSpreadsheetFormat,
	openCsv,
	openOds,
	openXlsb,
	openXlsx,
	XlsxError,
} from "@openjsxl/core";

/** One opener's verdict on a buffer: resolved, typed-rejected, or crashed (a real defect). */
export type Verdict = "ok" | "typed" | "crash";

export interface OpenerOutcome {
	readonly opener: string;
	readonly verdict: Verdict;
	/** The offending error, only when `verdict === "crash"`. */
	readonly error?: string;
}

export interface ReplayResult {
	readonly outcomes: readonly OpenerOutcome[];
	/** True when any opener crashed (threw a non-XlsxError). */
	readonly crashed: boolean;
}

// A generous decompressed-size cap: prevents a corrupted length field from making a campaign OOM,
// while leaving every real fixture part far under the ceiling. A part over the cap is a TYPED
// `part-too-large` rejection — still inside the invariant.
const MAX_PART_BYTES = 128 * 1024 * 1024;

// The minimum a workbook opener must expose so we can force cell materialization — structural, so any
// format's `Workbook` satisfies it without importing the concrete class.
interface Drainable {
	readonly sheets: readonly { readonly name: string }[];
	sheet(name: string): { rows(): AsyncGenerator<unknown> };
}

// Open then MATERIALIZE: drain every sheet's rows so the deferred sheet SAX parse actually runs.
async function openAndDrain(open: () => Drainable | Promise<Drainable>): Promise<void> {
	const book = await open();
	for (const info of book.sheets) {
		const ws = book.sheet(info.name);
		for await (const _row of ws.rows()) {
			// Pull every populated row — this is what triggers readRows / the cell parse.
		}
	}
}

function describe(e: unknown): string {
	if (e instanceof Error) return `${e.name}: ${e.message}`.slice(0, 200);
	return String(e).slice(0, 200);
}

async function runOpener(
	opener: string,
	fn: () => unknown | Promise<unknown>,
): Promise<OpenerOutcome> {
	try {
		await fn();
		return { opener, verdict: "ok" };
	} catch (e) {
		if (e instanceof XlsxError) return { opener, verdict: "typed" };
		return { opener, verdict: "crash", error: describe(e) };
	}
}

/**
 * Run `bytes` through all four openers (cells materialized) plus format detection, returning each
 * verdict. Every opener is tried regardless of the seed's original format — a reader handed the
 * "wrong" bytes must still reject typed, never crash.
 */
export async function replayAll(bytes: Uint8Array): Promise<ReplayResult> {
	const opts = { maxPartBytes: MAX_PART_BYTES } as const;
	const outcomes = await Promise.all([
		runOpener("openXlsx", () => openAndDrain(() => openXlsx(bytes, opts))),
		runOpener("openXlsb", () => openAndDrain(() => openXlsb(bytes, opts))),
		runOpener("openOds", () => openAndDrain(() => openOds(bytes, opts))),
		runOpener("openCsv", () => openAndDrain(() => openCsv(bytes))),
		runOpener("detect", () => detectSpreadsheetFormat(bytes)),
	]);
	return { outcomes, crashed: outcomes.some((o) => o.verdict === "crash") };
}
