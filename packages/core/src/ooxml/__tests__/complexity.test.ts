import { describe, expect, it } from "vitest";
import {
	bestDoublingRatio,
	doublingRatio,
	LINEAR_RATIO_LIMIT,
	MIN_SAMPLE_MS,
	timePerCall,
} from "./complexity";

describe("timePerCall", () => {
	it("repeats a fast call until a sample spans the minimum and reports per-call time", () => {
		let calls = 0;
		const perCall = timePerCall(() => {
			calls++;
		}, 1);
		expect(calls).toBeGreaterThan(1);
		expect(perCall).toBeGreaterThan(0);
		expect(perCall * calls).toBeGreaterThanOrEqual(MIN_SAMPLE_MS * 0.5);
	});

	it("rejects a non-positive sample count", () => {
		expect(() => timePerCall(() => 0, 0)).toThrow(RangeError);
	});

	it("runs a slow call once per sample", () => {
		let calls = 0;
		timePerCall(() => {
			calls++;
			const until = performance.now() + MIN_SAMPLE_MS + 1;
			while (performance.now() < until) {
				// busy-wait past the sample minimum
			}
		}, 3);
		expect(calls).toBe(3);
	});
});

describe("doublingRatio", () => {
	it("alternates small and large samples", () => {
		const order: number[] = [];
		const slow = (size: number) => () => {
			order.push(size);
			const until = performance.now() + MIN_SAMPLE_MS + 1;
			while (performance.now() < until) {
				// busy-wait past the sample minimum
			}
		};
		doublingRatio(slow, 1, 2);
		expect(order).toEqual([1, 2, 1, 2]);
	});

	it("builds the work for n and 2n and divides their times", () => {
		const sizes: number[] = [];
		const ratio = doublingRatio((size) => {
			sizes.push(size);
			return () => "x".repeat(size).split("");
		}, 20_000);
		expect(sizes).toEqual([20_000, 40_000]);
		expect(Number.isFinite(ratio)).toBe(true);
		expect(ratio).toBeGreaterThan(0);
	});
});

describe("bestDoublingRatio", () => {
	it("stops at the first attempt below the limit", () => {
		let builds = 0;
		const ratio = bestDoublingRatio((size) => {
			builds++;
			return () => "x".repeat(size).split("");
		}, 20_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
		expect(builds).toBe(2);
	});

	it("reports quadratic work above the limit after every attempt", () => {
		let builds = 0;
		const ratio = bestDoublingRatio(
			(size) => {
				builds++;
				return () => {
					let sum = 0;
					for (let i = 0; i < size * size; i++) sum += i;
					return sum;
				};
			},
			1_500,
			2,
		);
		expect(ratio).toBeGreaterThan(LINEAR_RATIO_LIMIT);
		expect(builds).toBe(4);
	});
});
