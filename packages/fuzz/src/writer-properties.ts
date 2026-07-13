// Half A — property-based fuzzing of the WRITER. Async fast-check properties, exported so both the CI
// smoke suite (small numRuns, fixed seed) and the long local run (large numRuns) share one definition.
//
// Corpus split (a soundness fix — the earlier all-hostile corpus rejected ~99.8% of inputs before any
// bytes were produced, so the resolve/re-read and determinism arms were near-vacuous):
//   • `validWorkbookArb`  — reliably WRITES; powers the determinism + resolve/re-read invariants and
//     asserts the writer never rejects legitimately-valid input.
//   • `hostileWorkbookArb` — poisoned values + hostile shapes; powers the never-crash invariant.
// P1 draws from BOTH (so it exercises resolve AND reject); P2/P4/P3 use the valid/scalar corpora.

import { openXlsx, type WorkbookInput, writeXlsx, XlsxError } from "@openjsxl/core";
import fc from "fast-check";
import { colToA1 } from "./a1";
import { hostileWorkbookArb, plainScalarWorkbookArb, validWorkbookArb } from "./arbitraries";

// The hostile corpus is intentionally malformed; the writer's contract here is to reject it with a
// typed error, never crash. This single boundary coercion confines the `unknown`→input step — core
// code stays cast-free (the whole point of Half A is feeding non-conforming input).
function writeUnknown(wb: unknown): Promise<Uint8Array> {
	return writeXlsx(wb as WorkbookInput);
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
	return true;
}

/**
 * P1 — resolve-or-typed-error. For ANY input (valid OR hostile: poisoned values, hostile shapes, wrong
 * types), `writeXlsx` must either resolve to bytes `openXlsx` can re-read, or throw
 * `XlsxError('invalid-input')`. A TypeError / RangeError / other code / a resolve whose bytes fail to
 * re-open is a defect (fast-check shrinks to a minimal reproducer). Drawing from both corpora means the
 * resolve+re-read arm actually fires (~half the runs) instead of being starved by rejections.
 */
export const resolveOrTypedError = fc.asyncProperty(
	fc.oneof(validWorkbookArb, hostileWorkbookArb),
	async (wb) => {
		let bytes: Uint8Array;
		try {
			bytes = await writeUnknown(wb);
		} catch (e) {
			if (e instanceof XlsxError && e.code === "invalid-input") return; // the one allowed failure
			throw e instanceof Error ? e : new Error(String(e));
		}
		await openXlsx(bytes); // it wrote — the bytes must be a real, re-readable .xlsx
	},
);

/**
 * P4 — valid input round-trips. A guaranteed-valid workbook (unique names, finite numbers, XML-safe
 * strings, legal styles, valid tables/DV/CF) MUST write (never a typed rejection of valid input) and
 * MUST re-open. This is where the "resolves to re-readable bytes" claim is exercised at ~100%.
 */
export const validRoundTrips = fc.asyncProperty(validWorkbookArb, async (wb) => {
	let bytes: Uint8Array;
	try {
		bytes = await writeXlsx(wb);
	} catch (e) {
		throw new Error(
			`writeXlsx rejected a VALID workbook: ${e instanceof Error ? e.message : String(e)}`,
		);
	}
	await openXlsx(bytes);
});

/**
 * P2 — deterministic bytes. Writing the SAME valid input twice yields byte-identical output. Guards the
 * "identical input → identical output" invariant (ordering / timestamp / iteration nondeterminism)
 * across styles, tables, DV, and CF. The valid corpus reliably writes, so the byte comparison actually
 * runs (it was starved to ~0% under the old plain corpus).
 */
export const deterministicBytes = fc.asyncProperty(validWorkbookArb, async (wb) => {
	const first = await writeXlsx(wb);
	const second = await writeXlsx(wb);
	if (!bytesEqual(first, second))
		throw new Error("nondeterministic writer output for identical input");
});

/**
 * P3 — value round-trip. A workbook of plain scalars writes AND reads back with every scalar preserved
 * (modulo the writer's documented type inference: `null` → empty). Catches silent corruption — a value
 * that survives write but comes back wrong.
 */
export const scalarRoundTrip = fc.asyncProperty(plainScalarWorkbookArb, async (wb) => {
	const bytes = await writeXlsx(wb);
	const book = await openXlsx(bytes);
	const sheet0 = wb.sheets[0];
	if (sheet0 === undefined) return;
	const ws = book.sheet(sheet0.name);
	sheet0.rows.forEach((row, r) => {
		if (row === undefined) return;
		row.forEach((expected, c) => {
			const cell = ws.cell(`${colToA1(c + 1)}${r + 1}`);
			if (expected === null) {
				if (cell.type !== "empty")
					throw new Error(`expected empty at r${r}c${c}, got ${cell.type}`);
				return;
			}
			if (cell.value !== expected) {
				throw new Error(
					`round-trip mismatch at r${r}c${c}: wrote ${JSON.stringify(expected)} read ${JSON.stringify(cell.value)} (${cell.type})`,
				);
			}
		});
	});
});
