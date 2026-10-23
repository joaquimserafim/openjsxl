import { describe, expect, it } from "vitest";
import {
	dateFormatDecider,
	isDateFormatCode,
	MAX_NUMBER_FORMAT_LEN,
	parseStyles,
	resolveNumberFormat,
	stripBracketSections,
} from "../styles";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "./complexity";
import {
	corpusFormatCodes,
	legacyIsDate,
	legacyStrip,
	randomCodes,
	stylesXml,
} from "./number-format-fixtures";

// F11.50: date-format detection is linear in code length and decided once per format id. The
// 1.0.0 scan (legacyStrip/legacyIsDate) is the equivalence oracle on short inputs.

describe("stripBracketSections", () => {
	it("removes bracketed sections and keeps an unclosed bracket with the rest", () => {
		expect(stripBracketSections("0.00")).toBe("0.00");
		expect(stripBracketSections("[Red]0.00")).toBe("0.00");
		expect(stripBracketSections("[h]:mm")).toBe(":mm");
		expect(stripBracketSections("0[x")).toBe("0[x");
		expect(stripBracketSections("[a]b[c")).toBe("b[c");
	});

	it("matches the 1.0.0 pass on crafted, corpus and random codes", async () => {
		const crafted = [
			"",
			"[",
			"[[[",
			"]]",
			"]x[",
			"[[x]",
			"[a[b]c]",
			"[a][b][",
			"[[]]",
			"[>100]0;[<0]0",
			"[$-409]mmmm d, yyyy",
		];
		const codes = [...crafted, ...(await corpusFormatCodes()), ...randomCodes(3000)];
		for (const code of codes) expect(stripBracketSections(code), code).toBe(legacyStrip(code));
	});

	it("is linear in code length on unclosed brackets", () => {
		const ratio = bestDoublingRatio((n) => () => stripBracketSections("[".repeat(n)), 25_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 60_000);
});

describe("isDateFormatCode", () => {
	it("keeps every corpus and random code's date decision", async () => {
		const corpus = await corpusFormatCodes();
		expect(corpus.length).toBeGreaterThan(0);
		for (const code of [...corpus, ...randomCodes(3000)]) {
			expect(isDateFormatCode(code), code).toBe(legacyIsDate(code));
		}
	});

	it("keeps every corpus code within Excel's number-format length", async () => {
		for (const code of await corpusFormatCodes()) {
			expect(code.length, code).toBeLessThanOrEqual(MAX_NUMBER_FORMAT_LEN);
		}
	});

	// Both shapes are quadratic in 1.0.0 (ratio about 4). Shapes that were already linear add only
	// timing noise, so they are not pinned here.
	it.each([
		["unclosed brackets", (n: number) => "[".repeat(n)],
		["elapsed-time openers", (n: number) => "[h".repeat(n / 2)],
	])(
		"is linear in code length on %s",
		(_name, shape) => {
			const ratio = bestDoublingRatio((n) => () => isDateFormatCode(shape(n)), 25_000);
			expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
		},
		60_000,
	);
});

describe("MAX_NUMBER_FORMAT_LEN", () => {
	it("is Excel's 255-character limit", () => {
		expect(MAX_NUMBER_FORMAT_LEN).toBe(255);
	});
});

describe("resolveNumberFormat", () => {
	it("prefers a custom code, keeps a dropped code absent and falls back to the built-in table", () => {
		const custom = new Map<number, string | null>([
			[164, "0.0"],
			[14, null],
		]);
		expect(resolveNumberFormat(custom, 164)).toBe("0.0");
		expect(resolveNumberFormat(custom, 14)).toBeUndefined();
		expect(resolveNumberFormat(custom, 2)).toBe("0.00");
		expect(resolveNumberFormat(custom, 300)).toBeUndefined();
	});
});

describe("dateFormatDecider", () => {
	it("decides each custom id once and serves later calls from the cache", () => {
		let calls = 0;
		const isDateId = dateFormatDecider(
			new Map([
				[164, "yyyy-mm-dd"],
				[165, "0.00"],
			]),
			(code) => {
				calls++;
				return isDateFormatCode(code);
			},
		);
		expect([isDateId(164), isDateId(164), isDateId(165), isDateId(165)]).toEqual([
			true,
			true,
			false,
			false,
		]);
		expect(calls).toBe(2);
	});

	it("uses the built-in table for ids without a custom code", () => {
		let calls = 0;
		const isDateId = dateFormatDecider(new Map(), () => {
			calls++;
			return true;
		});
		expect([isDateId(14), isDateId(22), isDateId(1), isDateId(49)]).toEqual([
			true,
			true,
			false,
			false,
		]);
		expect(calls).toBe(0);
	});

	it("treats a dropped custom code as no date format, without the built-in fallback", () => {
		const isDateId = dateFormatDecider(new Map([[14, null]]));
		expect(isDateId(14)).toBe(false);
	});
});

describe("parseStyles date decisions", () => {
	it("decides each custom format id once however many cells and xfs use it", () => {
		let calls = 0;
		const styles = parseStyles(
			stylesXml(
				[
					[164, "yyyy-mm-dd"],
					[165, "0.0"],
				],
				[0, 164, 165, 164, 165],
			),
			(code) => {
				calls++;
				return isDateFormatCode(code);
			},
		);
		const decisions: boolean[] = [];
		for (let cell = 0; cell < 100; cell++) decisions.push(styles.isDateStyle(1 + (cell % 4)));
		expect(decisions.slice(0, 4)).toEqual([true, false, true, false]);
		expect(calls).toBe(2);
	});
});
