import { describe, expect, it } from "vitest";
import { expectedReadSink, validateReadResult, validateReadResults } from "../read-validation.mjs";

const read = { ok: true, op: "read", workload: "numbers", rows: 2, sink: 291.3 };

describe("read benchmark validation", () => {
	it("derives numeric, styled, and string sinks from the known workload values", () => {
		expect(expectedReadSink("numbers", 1)).toBeCloseTo(70.65, 10);
		expect(expectedReadSink("numbers", 2)).toBeCloseTo(291.3, 10);
		expect(expectedReadSink("styled", 2)).toBe(expectedReadSink("numbers", 2));
		expect(expectedReadSink("strings", 2)).toBe(144);
		expect(expectedReadSink("numbers", 0)).toBe(0);
	});

	it("accepts finite correct sums and small floating-point rounding differences", () => {
		const result = validateReadResult(read);
		expect(result.ok).toBe(true);
		expect(result.expectedSink).toBeCloseTo(291.3, 10);
		expect(result.actualSink).toBe(291.3);
		const expected = expectedReadSink("numbers", 1_000);
		expect(
			validateReadResult({
				...read,
				rows: 1_000,
				sink: expected + expected * Number.EPSILON * 16,
			}).ok,
		).toBe(true);
		expect(
			validateReadResult({
				...read,
				rows: 1_000,
				sink: expected + expected * Number.EPSILON * 64,
			}).ok,
		).toBe(false);
	});

	it("rejects the first-column-only XLSB result and retains expected/actual diagnostics", () => {
		const result = validateReadResult({
			...read,
			format: "xlsb",
			rows: 1_000,
			sink: 7_492_500,
		});
		expect(result.ok).toBe(false);
		expect(result.reason).toContain("Read checksum mismatch");
		expect(result.expectedSink).toBeCloseTo(74_995_650, 4);
		expect(result.actualSink).toBe(7_492_500);
		expect(JSON.parse(JSON.stringify(result)).actualSink).toBe(7_492_500);
	});

	it("rejects nonfinite, missing, and nonnumeric sinks, including JSON's null conversion", () => {
		for (const sink of [NaN, Infinity, -Infinity, undefined, null, "291.3"]) {
			const result = validateReadResult({ ...read, sink });
			expect(result.ok).toBe(false);
			expect(result.reason).toContain("Read checksum mismatch");
			expect(JSON.parse(JSON.stringify(result)).actualSink).toBe(String(sink));
		}
	});

	it("leaves write measurements and previously failed cells unchanged", () => {
		const write = { ok: true, op: "write", outBytes: 123 };
		const unsupported = { ...read, ok: false, reason: "unsupported" };
		const results = validateReadResults([read, write, unsupported]);
		expect(results[0].ok).toBe(true);
		expect(results[1]).toBe(write);
		expect(results[2]).toBe(unsupported);
	});

	it("rejects missing or invalid workload metadata rather than reporting an unchecked read", () => {
		for (const change of [
			{ rows: undefined },
			{ rows: -1 },
			{ rows: 1.5 },
			{ workload: "unknown" },
		]) {
			expect(validateReadResult({ ...read, ...change }).ok).toBe(false);
		}
	});
});
