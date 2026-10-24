/** A sample this long (ms) is large enough that timer resolution cannot dominate a ratio. */
export const MIN_SAMPLE_MS = 20;

/**
 * The calling thread's CPU time in milliseconds, user plus system: the helpers' default clock
 * (F11.94). Unlike wall time it does not advance while other jobs hold the CPU or while the thread
 * sleeps, so a pin measures its own work. Work on other threads (concurrent GC, inflate) is not
 * counted, so it suits synchronous work only; an async pin would need another clock.
 *
 * @example threadCpuMs() // e.g. 412.38: grows only while this thread runs
 */
export function threadCpuMs(): number {
	const { user, system } = process.threadCpuUsage();
	return (user + system) / 1000;
}

/**
 * Median time of one `fn()` call over `samples` samples, on the clock `now` (milliseconds; the
 * calling thread's CPU time unless given, and the self-tests pass a fake one). A call faster than
 * `MIN_SAMPLE_MS` is repeated (doubling the count) until one sample spans at least that long; a slow
 * call runs once.
 *
 * @example timePerCall(() => "x".repeat(1000)) // a small positive number of milliseconds
 * @example let t = 0; timePerCall(() => { t += 25; }, 1, () => t) // 25
 * @example timePerCall(() => 0, 0) // throws RangeError: samples must be a positive integer
 */
export function timePerCall(
	fn: () => unknown,
	samples = 3,
	now: () => number = threadCpuMs,
): number {
	if (!Number.isInteger(samples) || samples < 1)
		throw new RangeError("samples must be a positive integer");
	const perCall: number[] = [];
	for (let s = 0; s < samples; s++) {
		let reps = 1;
		for (;;) {
			const start = now();
			for (let i = 0; i < reps; i++) fn();
			const elapsed = now() - start;
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
 * slow sample (noise only ever adds time) cannot inflate one side alone. `now` is the clock, the
 * calling thread's CPU time unless given.
 *
 * @example doublingRatio((n) => () => "x".repeat(n).split(""), 50_000) // about 2
 * @example let t = 0; doublingRatio((n) => () => { t += n; }, 1_000, 3, () => t) // exactly 2
 */
export function doublingRatio(
	make: (size: number) => () => unknown,
	n: number,
	samples = 3,
	now: () => number = threadCpuMs,
): number {
	const small = make(n);
	const large = make(2 * n);
	let fastestSmall = Number.POSITIVE_INFINITY;
	let fastestLarge = Number.POSITIVE_INFINITY;
	for (let s = 0; s < samples; s++) {
		fastestSmall = Math.min(fastestSmall, timePerCall(small, 1, now));
		fastestLarge = Math.min(fastestLarge, timePerCall(large, 1, now));
	}
	return fastestLarge / fastestSmall;
}

/** A complexity pin's ratio limit: linear work measures about 2, quadratic about 4. */
export const LINEAR_RATIO_LIMIT = 3;

/**
 * The lowest `doublingRatio` over up to `attempts` measurements, stopping at the first one below
 * `LINEAR_RATIO_LIMIT`. Retrying protects a linear pin from a noisy measurement: noise can push one
 * measurement over the limit, but not every attempt. It does not make quadratic work measure 4:
 * noise on the small-n samples lowers a quadratic ratio, which is why the default clock is the
 * thread's CPU time rather than wall time, which also counts other jobs. `now` is the clock.
 *
 * @example bestDoublingRatio((n) => () => "x".repeat(n).split(""), 50_000) // below 3 (linear)
 * @example bestDoublingRatio((n) => () => { for (let i = 0; i < n * n; i++); }, 2_000) // about 4
 * @example let t = 0; bestDoublingRatio((n) => () => { t += (n * n) / 1_000; }, 1_000, 2, () => t) // exactly 4, after both attempts
 */
export function bestDoublingRatio(
	make: (size: number) => () => unknown,
	n: number,
	attempts = 3,
	now: () => number = threadCpuMs,
): number {
	let best = Number.POSITIVE_INFINITY;
	for (let a = 0; a < attempts && best >= LINEAR_RATIO_LIMIT; a++) {
		best = Math.min(best, doublingRatio(make, n, undefined, now));
	}
	return best;
}
