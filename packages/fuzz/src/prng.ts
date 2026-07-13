// A tiny seeded PRNG (xorshift32) — zero dependencies, fully deterministic, so a crasher is
// reproducible from its seed alone. NOT cryptographic; the point is a repeatable byte stream for the
// mutation engine (Half B) and for choosing mutators. fast-check owns its own randomness for Half A.

/** A deterministic 32-bit xorshift generator seeded by a non-zero integer. */
export class Prng {
	#state: number;

	constructor(seed: number) {
		// xorshift jams on a zero state; fold the seed into a guaranteed-non-zero 32-bit word.
		const s = (seed | 0) ^ 0x9e3779b9;
		this.#state = s === 0 ? 0x1234_5678 : s >>> 0;
	}

	/** Next raw 32-bit unsigned word. */
	next(): number {
		let x = this.#state;
		x ^= x << 13;
		x ^= x >>> 17;
		x ^= x << 5;
		this.#state = x >>> 0;
		return this.#state;
	}

	/** A float in [0, 1). */
	float(): number {
		return this.next() / 0x1_0000_0000;
	}

	/** An integer in [0, n) (n treated as a positive, safe integer). */
	int(n: number): number {
		if (n <= 1) return 0;
		return Math.floor(this.float() * n);
	}

	/** A byte value 0–255. */
	byte(): number {
		return this.next() & 0xff;
	}

	/** Pick one element of a non-empty array. */
	pick<T>(items: readonly T[]): T {
		// Callers pass non-empty arrays; the modulo guard keeps the index in range regardless.
		const i = this.int(items.length);
		return items[i] as T;
	}
}
