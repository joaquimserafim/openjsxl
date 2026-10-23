/** A sample this long (ms) is large enough that timer resolution cannot dominate a ratio. */
export const MIN_SAMPLE_MS = 20;

/**
 * Median wall time of one `fn()` call over `samples` samples. A call faster than `MIN_SAMPLE_MS` is
 * repeated (doubling the count) until one sample spans at least that long; a slow call runs once.
 *
 * @example timePerCall(() => "x".repeat(1000)) // a small positive number of milliseconds
 * @example timePerCall(() => 0, 0) // throws RangeError: samples must be a positive integer
 */
export function timePerCall(fn: () => unknown, samples = 3): number {
	if (!Number.isInteger(samples) || samples < 1)
		throw new RangeError("samples must be a positive integer");
	const perCall: number[] = [];
	for (let s = 0; s < samples; s++) {
		let reps = 1;
		for (;;) {
			const start = performance.now();
			for (let i = 0; i < reps; i++) fn();
			const elapsed = performance.now() - start;
			if (elapsed >= MIN_SAMPLE_MS) {
				perCall.push(elapsed / reps);
				break;
			}
			reps *= 2;
		}
	}
	perCall.sort((a, b) => a - b);
	return perCall[Math.floor(perCall.length / 2)] as number;
}

/**
 * `time(2n) / time(n)` for the work `make(size)` returns: about 2 for linear work and about 4 for
 * quadratic work. Small and large samples alternate and each size keeps its fastest sample, so a
 * load spike (which only ever adds time) cannot inflate one side alone.
 *
 * @example doublingRatio((n) => () => "x".repeat(n).split(""), 50_000) // about 2
 */
export function doublingRatio(
	make: (size: number) => () => unknown,
	n: number,
	samples = 3,
): number {
	const small = make(n);
	const large = make(2 * n);
	let fastestSmall = Number.POSITIVE_INFINITY;
	let fastestLarge = Number.POSITIVE_INFINITY;
	for (let s = 0; s < samples; s++) {
		fastestSmall = Math.min(fastestSmall, timePerCall(small, 1));
		fastestLarge = Math.min(fastestLarge, timePerCall(large, 1));
	}
	return fastestLarge / fastestSmall;
}

/** A complexity pin's ratio limit: linear work measures about 2, quadratic about 4. */
export const LINEAR_RATIO_LIMIT = 3;

/**
 * The lowest `doublingRatio` over up to `attempts` measurements, stopping at the first one below
 * `LINEAR_RATIO_LIMIT`. Linear work passes on the first try on an idle machine; quadratic work never
 * measures below about 4, so retrying cannot hide it.
 *
 * @example bestDoublingRatio((n) => () => "x".repeat(n).split(""), 50_000) // below 3 (linear)
 * @example bestDoublingRatio((n) => () => { for (let i = 0; i < n * n; i++); }, 2_000) // about 4
 */
export function bestDoublingRatio(
	make: (size: number) => () => unknown,
	n: number,
	attempts = 3,
): number {
	let best = Number.POSITIVE_INFINITY;
	for (let a = 0; a < attempts && best >= LINEAR_RATIO_LIMIT; a++) {
		best = Math.min(best, doublingRatio(make, n));
	}
	return best;
}
