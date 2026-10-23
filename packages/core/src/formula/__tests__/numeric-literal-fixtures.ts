// F11.51 fixtures: the 1.0.0 numeric-literal check, kept as an equivalence oracle, and an exhaustive
// enumerator of short strings.

/** The 1.0.0 en-US literal pattern. A long digit run that fails to match backtracks quadratically. */
const LEGACY_EN_US_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * The 1.0.0 `numericStringValue`, kept only as an oracle for short inputs.
 *
 * @example legacyNumericStringValue(" 1.5 ") // 1.5
 * @example legacyNumericStringValue("1..2") // undefined
 */
export function legacyNumericStringValue(s: string): number | undefined {
	const t = s.trim();
	if (t !== "" && LEGACY_EN_US_NUMBER.test(t)) {
		const n = Number(t);
		if (Number.isFinite(n)) return n;
	}
	return undefined;
}

/**
 * Every string over the characters of `alphabet` with length 0 to `maxLength`, shortest first.
 *
 * @example [...allStrings("ab", 2)] // ["", "a", "b", "aa", "ab", "ba", "bb"]
 */
export function* allStrings(alphabet: string, maxLength: number): Generator<string> {
	let level = [""];
	yield "";
	for (let length = 1; length <= maxLength; length++) {
		const next: string[] = [];
		for (const prefix of level) for (const ch of alphabet) next.push(prefix + ch);
		yield* next;
		level = next;
	}
}
