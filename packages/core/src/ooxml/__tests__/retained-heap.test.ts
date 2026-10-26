import { describe, expect, it } from "vitest";
import { collectGarbage, retainedHeapBytes } from "./retained-heap";

const MB = 2 ** 20;

describe("collectGarbage", () => {
	it("frees what is no longer reachable", () => {
		collectGarbage();
		const before = process.memoryUsage().heapUsed;
		let garbage: number[] | undefined = new Array(4_000_000).fill(1);
		expect(garbage.length).toBe(4_000_000);
		garbage = undefined;
		collectGarbage();
		expect(process.memoryUsage().heapUsed - before).toBeLessThan(4 * MB);
	});
});

describe("retainedHeapBytes", () => {
	it("counts what the result keeps alive", () => {
		// 4M small integers in a packed array: 8 bytes each.
		const retained = retainedHeapBytes(() => new Array(4_000_000).fill(1));
		expect(retained).toBeGreaterThan(24 * MB);
		expect(retained).toBeLessThan(40 * MB);
	});

	it("does not count what the build drops", () => {
		const retained = retainedHeapBytes(() => new Array(4_000_000).fill(1).length);
		expect(retained).toBeLessThan(4 * MB);
	});

	it("tells a string built a character at a time from a flat one", () => {
		const flat = retainedHeapBytes(() => Array.from({ length: 1_000_000 }, () => "x").join(""));
		const rope = retainedHeapBytes(() => {
			let s = "";
			for (let i = 0; i < 1_000_000; i++) s += "x";
			return s;
		});
		expect(flat).toBeLessThan(4 * MB);
		expect(rope).toBeGreaterThan(16 * MB);
	});
});
