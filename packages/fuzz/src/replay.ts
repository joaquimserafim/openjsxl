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
//
// Each ZIP opener also runs with `verifyChecksums: true`, and `streamSheetRows` streams every sheet
// that the buffered open loaded, plain and checked. The stream must give the same rows as the buffered
// `rows()`, or fail typed when they do (F11.47), so the CRC path and the streaming end-of-entry checks
// are fuzzed too.

import {
	detectSpreadsheetFormat,
	openCsv,
	openOds,
	openXlsb,
	openXlsx,
	type ReadOptions,
	streamSheetRows,
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
const PLAIN: ReadOptions = { maxPartBytes: MAX_PART_BYTES };
const CHECKED: ReadOptions = { maxPartBytes: MAX_PART_BYTES, verifyChecksums: true };

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

/**
 * Every row a row stream yields, as JSON text, or `"typed"` when it fails with an
 * {@link XlsxError}. Any other failure is rethrown, so the lane reports it as a crash.
 *
 * @example await rowsOrTyped(book.sheet("S").rows()) // '[{"index":1,"cells":[…]}]'
 * @example await rowsOrTyped(streamSheetRows(cut, "S")) // "typed"
 */
export async function rowsOrTyped(rows: AsyncIterable<unknown>): Promise<string> {
	const read = async (): Promise<string> => {
		const out: unknown[] = [];
		for await (const row of rows) out.push(row);
		return JSON.stringify(out);
	};
	return read().catch((e: unknown) => {
		if (e instanceof XlsxError) return "typed";
		throw e;
	});
}

/**
 * Two reads of one sheet, as {@link rowsOrTyped} gives them, must be the same text. A disagreement
 * throws a plain `Error`, which the lane reports as a crash.
 *
 * @example assertSameRows("S", "typed", "typed") // returns
 * @example assertSameRows("S", "[]", "typed") // throws Error "streamSheetRows and openXlsx disagree on sheet \"S\": …"
 */
export function assertSameRows(sheet: string, buffered: string, streamed: string): void {
	if (streamed === buffered) return;
	throw new Error(
		`streamSheetRows and openXlsx disagree on sheet ${JSON.stringify(sheet)}: ${streamed.slice(0, 60)} vs ${buffered.slice(0, 60)}`,
	);
}

/**
 * Open `bytes` buffered, then stream each sheet it loaded with the same options. Both reads of a sheet
 * must give the same rows, or both must fail typed ({@link assertSameRows}). When the buffered open
 * fails typed, the first sheet is still streamed, as a no-crash check of the stream's own size and
 * CRC-32 checks on the damaged entries, and the open's error is rethrown.
 *
 * @example await streamMatchesBuffered(await loadFixture("basic.xlsx"), {}) // resolves
 * @example await streamMatchesBuffered(crcDamaged, { verifyChecksums: true }) // streams the first sheet, then rejects XlsxError "corrupt-zip"
 */
export async function streamMatchesBuffered(
	bytes: Uint8Array,
	options: ReadOptions,
): Promise<void> {
	const opened = await openXlsx(bytes, options).then(
		(book) => ({ book }),
		(error: unknown) => {
			if (error instanceof XlsxError) return { error };
			throw error;
		},
	);
	if ("error" in opened) {
		await rowsOrTyped(streamSheetRows(bytes, undefined, options));
		throw opened.error;
	}
	const { book } = opened;
	for (const info of book.sheets) {
		const buffered = await rowsOrTyped(book.sheet(info.name).rows());
		const streamed = await rowsOrTyped(streamSheetRows(bytes, info.name, options));
		assertSameRows(info.name, buffered, streamed);
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
 * "wrong" bytes must still reject typed, never crash. The ZIP openers run again with
 * `verifyChecksums: true`, and two `streamSheetRows` lanes compare the stream with the buffered rows.
 */
export async function replayAll(bytes: Uint8Array): Promise<ReplayResult> {
	const outcomes = await Promise.all([
		runOpener("openXlsx", () => openAndDrain(() => openXlsx(bytes, PLAIN))),
		runOpener("openXlsb", () => openAndDrain(() => openXlsb(bytes, PLAIN))),
		runOpener("openOds", () => openAndDrain(() => openOds(bytes, PLAIN))),
		runOpener("openCsv", () => openAndDrain(() => openCsv(bytes))),
		runOpener("detect", () => detectSpreadsheetFormat(bytes)),
		runOpener("openXlsx+crc", () => openAndDrain(() => openXlsx(bytes, CHECKED))),
		runOpener("openXlsb+crc", () => openAndDrain(() => openXlsb(bytes, CHECKED))),
		runOpener("openOds+crc", () => openAndDrain(() => openOds(bytes, CHECKED))),
		runOpener("streamSheetRows", () => streamMatchesBuffered(bytes, PLAIN)),
		runOpener("streamSheetRows+crc", () => streamMatchesBuffered(bytes, CHECKED)),
	]);
	return { outcomes, crashed: outcomes.some((o) => o.verdict === "crash") };
}
