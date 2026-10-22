import { rowAt, WORKLOADS } from "./workloads.mjs";

const expectedSinks = new Map();

// Reuse the deterministic workload recipe without allocating a whole workbook. This runs in the
// orchestrator after a worker exits (or while re-rendering), never in a timed operation.
export function expectedReadSink(kind, rows) {
	if (!WORKLOADS.includes(kind) || !Number.isSafeInteger(rows) || rows < 0) {
		throw new Error("invalid workload or row count for read checksum");
	}
	const key = `${kind}:${rows}`;
	if (expectedSinks.has(key)) return expectedSinks.get(key);
	let sink = 0;
	for (let r = 0; r < rows; r++) {
		for (const cell of rowAt(kind, r)) {
			const value = kind === "styled" ? cell.value : cell;
			sink += typeof value === "number" ? value : value.length;
		}
	}
	expectedSinks.set(key, sink);
	return sink;
}

export function validateReadResult(result) {
	if (result.op !== "read" || !result.ok) return result;
	const actual = result.sink;
	// JSON turns NaN/Infinity into null; keep a readable diagnostic for direct/helper callers too.
	const actualSink = Number.isFinite(actual) ? actual : String(actual);
	let expectedSink;
	try {
		expectedSink = expectedReadSink(result.workload, result.rows);
	} catch (error) {
		return {
			...result,
			ok: false,
			actualSink,
			reason: `Read checksum validation failed: ${error.message}`,
		};
	}
	const checksumTolerance = Math.max(1e-9, Math.abs(expectedSink) * Number.EPSILON * 32);
	const valid = Number.isFinite(actual) && Math.abs(actual - expectedSink) <= checksumTolerance;
	return {
		...result,
		ok: valid,
		readValidationVersion: 1,
		expectedSink,
		actualSink,
		checksumTolerance,
		...(valid
			? {}
			: {
					reason: `Read checksum mismatch: expected ${expectedSink}, received ${actualSink} (tolerance ${checksumTolerance})`,
				}),
	};
}

export function validateReadResults(results) {
	return results.map(validateReadResult);
}
