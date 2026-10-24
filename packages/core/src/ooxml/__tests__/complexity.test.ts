import { afterEach, beforeEach, describe, expect, it, type MockInstance, vi } from "vitest";
import {
	bestDoublingRatio,
	doublingRatio,
	LINEAR_RATIO_LIMIT,
	MIN_SAMPLE_MS,
	threadCpuMs,
	timePerCall,
} from "./complexity";

// F11.94: the helpers' self-tests drive them with a fake clock that each work function advances by a
// known cost, so every ratio and count is exact and no self-test reads a real clock. The product pins
// use the default clock, the calling thread's CPU time, which `threadCpuMs` pins on its own.

let wallClock: MockInstance<() => number>;

beforeEach(() => {
	wallClock = vi.spyOn(performance, "now");
});

afterEach(() => {
	expect(wallClock).not.toHaveBeenCalled();
	wallClock.mockRestore();
});

describe("threadCpuMs", () => {
	it("reads the thread's user plus system CPU time, in milliseconds", () => {
		const usage = vi
			.spyOn(process, "threadCpuUsage")
			.mockReturnValueOnce({ user: 1_500, system: 500 });
		expect(threadCpuMs()).toBe(2);
		usage.mockRestore();
	});

	it("barely advances while the thread sleeps", async () => {
		const before = threadCpuMs();
		await new Promise((resolve) => setTimeout(resolve, 50));
		expect(threadCpuMs() - before).toBeLessThan(10);
	});

	it("advances during CPU work", () => {
		const before = threadCpuMs();
		let sum = 0;
		for (let i = 0; i < 30_000_000; i++) sum += i % 7;
		expect(sum).toBeGreaterThan(0);
		expect(threadCpuMs() - before).toBeGreaterThan(5);
	});
});

describe("timing helpers on a fake clock", () => {
	let cpuClock: MockInstance<typeof process.threadCpuUsage>;

	beforeEach(() => {
		cpuClock = vi.spyOn(process, "threadCpuUsage");
	});

	afterEach(() => {
		expect(cpuClock).not.toHaveBeenCalled();
		cpuClock.mockRestore();
	});

	describe("timePerCall", () => {
		it("repeats a fast call until a sample spans the minimum and reports per-call time", () => {
			let t = 0;
			let calls = 0;
			const perCall = timePerCall(
				() => {
					calls++;
					t += 1;
				},
				1,
				() => t,
			);
			// 1 ms per call: 1, 2, 4, 8 and 16 calls stay under 20 ms; 32 calls span 32 ms.
			expect(MIN_SAMPLE_MS).toBe(20);
			expect(calls).toBe(1 + 2 + 4 + 8 + 16 + 32);
			expect(perCall).toBe(1);
		});

		it("rejects a non-positive sample count", () => {
			expect(() =>
				timePerCall(
					() => 0,
					0,
					() => 0,
				),
			).toThrow(RangeError);
		});

		it("runs a slow call once per sample and reports the median sample", () => {
			let t = 0;
			const costs = [30, 50, 40];
			let calls = 0;
			const perCall = timePerCall(
				() => {
					t += costs[calls++ % costs.length] as number;
				},
				3,
				() => t,
			);
			expect(calls).toBe(3);
			expect(perCall).toBe(40);
		});
	});

	describe("doublingRatio", () => {
		it("alternates small and large samples", () => {
			let t = 0;
			const order: number[] = [];
			doublingRatio(
				(size) => () => {
					order.push(size);
					t += MIN_SAMPLE_MS + 1;
				},
				1,
				2,
				() => t,
			);
			expect(order).toEqual([1, 2, 1, 2]);
		});

		it("measures exactly 2 for linear work and 4 for quadratic work", () => {
			let t = 0;
			const now = () => t;
			const linear = doublingRatio(
				(size) => () => {
					t += size;
				},
				1_000,
				3,
				now,
			);
			const quadratic = doublingRatio(
				(size) => () => {
					t += (size * size) / 1_000;
				},
				1_000,
				3,
				now,
			);
			expect([linear, quadratic]).toEqual([2, 4]);
		});

		it("keeps each size's fastest sample, so a slow sample on one side does not move the ratio", () => {
			let t = 0;
			const ratio = doublingRatio(
				(size) => {
					const slowdowns = size === 1_000 ? [1, 1, 1] : [5, 1, 3];
					let call = 0;
					return () => {
						t += size * (slowdowns[call++ % slowdowns.length] as number);
					};
				},
				1_000,
				3,
				() => t,
			);
			expect(ratio).toBe(2);
		});

		it("builds the work for n and 2n once each", () => {
			let t = 0;
			const sizes: number[] = [];
			doublingRatio(
				(size) => {
					sizes.push(size);
					return () => {
						t += size;
					};
				},
				20_000,
				3,
				() => t,
			);
			expect(sizes).toEqual([20_000, 40_000]);
		});
	});

	describe("bestDoublingRatio", () => {
		it("stops at the first attempt below the limit", () => {
			let t = 0;
			let builds = 0;
			const ratio = bestDoublingRatio(
				(size) => {
					builds++;
					return () => {
						t += size;
					};
				},
				1_000,
				3,
				() => t,
			);
			expect(ratio).toBe(2);
			expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
			expect(builds).toBe(2);
		});

		it("reports quadratic work above the limit after every attempt", () => {
			let t = 0;
			let builds = 0;
			const ratio = bestDoublingRatio(
				(size) => {
					builds++;
					return () => {
						t += (size * size) / 1_000;
					};
				},
				1_000,
				2,
				() => t,
			);
			expect(ratio).toBe(4);
			expect(builds).toBe(4);
		});

		it("keeps the lowest ratio across attempts", () => {
			let t = 0;
			let builds = 0;
			const ratio = bestDoublingRatio(
				(size) => {
					builds++;
					// The first attempt measures quadratic work, the second linear work.
					const quadratic = builds <= 2;
					return () => {
						t += quadratic ? (size * size) / 1_000 : size;
					};
				},
				1_000,
				3,
				() => t,
			);
			expect(ratio).toBe(2);
			expect(builds).toBe(4);
		});
	});
});
