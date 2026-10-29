import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import {
	compileGlob,
	findSegment,
	type GlobSegment,
	type GlobToken,
	globMatch,
	globSegment,
	segmentAt,
} from "../builtins";

// F11.68: wildcard patterns compile to their runs between `*`s, matched greedily, leftmost first. A run
// with a `?` is found with a Shift-And scan (D12). Every result is the one the two-pointer matcher
// before F11.68 gave, copied below as the reference.

function referenceMatch(text: string, pattern: string): boolean {
	const tokens: GlobToken[] = [];
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern.charAt(i);
		if (ch === "~") {
			const next = pattern.charAt(i + 1);
			if (next === "*" || next === "?" || next === "~") {
				tokens.push({ kind: "lit", ch: next.toUpperCase() });
				i++;
				continue;
			}
			tokens.push({ kind: "lit", ch: "~" });
			continue;
		}
		if (ch === "*") tokens.push({ kind: "any" });
		else if (ch === "?") tokens.push({ kind: "one" });
		else tokens.push({ kind: "lit", ch: ch.toUpperCase() });
	}
	const up = text.toUpperCase();
	let i = 0;
	let j = 0;
	let star = -1;
	let mark = 0;
	while (i < up.length) {
		const t = tokens[j];
		if (t !== undefined && (t.kind === "one" || (t.kind === "lit" && t.ch === up.charAt(i)))) {
			i++;
			j++;
		} else if (t !== undefined && t.kind === "any") {
			star = j;
			mark = i;
			j++;
		} else if (star !== -1) {
			j = star + 1;
			mark++;
			i = mark;
		} else {
			return false;
		}
	}
	while (j < tokens.length && tokens[j]?.kind === "any") j++;
	return j === tokens.length;
}

const match = (text: string, pattern: string) => globMatch(text, compileGlob(pattern));
const run = (pattern: string) => compileGlob(pattern).segments[0] as GlobSegment;

describe("globSegment", () => {
	it("keeps a run with no ? as one needle", () => {
		const segment = globSegment([
			{ kind: "lit", ch: "A" },
			{ kind: "lit", ch: "B" },
		]);
		expect(segment.needle).toBe("AB");
		expect(segment.chars).toEqual(["A", "B"]);
	});

	it("gives a run with a ? one Shift-And mask per character", () => {
		const segment = globSegment([
			{ kind: "lit", ch: "A" },
			{ kind: "one" },
			{ kind: "lit", ch: "A" },
		]);
		expect(segment.needle).toBeUndefined();
		expect([...segment.rest]).toEqual([0b010]);
		expect([...(segment.masks.get(65) ?? [])]).toEqual([0b111]);
	});

	it("keeps a run of 32 characters as a needle, and gives a longer one masks", () => {
		expect(run("a".repeat(32)).needle).toBe("A".repeat(32));
		const long = run("a".repeat(33));
		expect(long.needle).toBeUndefined();
		expect([...long.rest]).toEqual([0, 0]);
		expect([...(long.masks.get(65) ?? [])]).toEqual([0xffffffff, 1]);
	});

	it("spreads a run longer than 32 over more words", () => {
		const segment = run(`${"?".repeat(32)}a`);
		expect([...segment.rest]).toEqual([0xffffffff, 0]);
		expect([...(segment.masks.get(65) ?? [])]).toEqual([0xffffffff, 1]);
	});
});

describe("compileGlob", () => {
	it("splits the pattern at each *", () => {
		expect(compileGlob("a*b?*c").segments.map((s) => s.chars)).toEqual([
			["A"],
			["B", undefined],
			["C"],
		]);
		expect(compileGlob("*").segments.map((s) => s.chars)).toEqual([[], []]);
		expect(compileGlob("ab").segments.map((s) => s.chars)).toEqual([["A", "B"]]);
	});

	it("reads ~ as an escape for *, ? and ~ only", () => {
		expect(compileGlob("~*~?~~~a").segments.map((s) => s.chars)).toEqual([
			["*", "?", "~", "~", "A"],
		]);
	});

	it("marks a pattern that can never match, where a literal upper-cases to two characters", () => {
		expect(compileGlob("*ß*").never).toBe(true);
		expect(compileGlob("*ﬁ*").never).toBe(true);
		expect(compileGlob("*s*").never).toBe(false);
	});
});

describe("segmentAt", () => {
	it("checks a run at one position, inside the text", () => {
		expect(segmentAt(run("b?"), "ABC", 1)).toBe(true);
		expect(segmentAt(run("b?"), "ABC", 0)).toBe(false);
		expect(segmentAt(run("b?"), "ABC", 2)).toBe(false);
		expect(segmentAt(run("b?"), "ABC", -1)).toBe(false);
		expect(segmentAt(run("b?"), "AB", 1)).toBe(false);
		expect(segmentAt(run(""), "", 0)).toBe(true);
	});
});

describe("findSegment", () => {
	it("finds a run with no ? with indexOf", () => {
		expect(findSegment(run("bc"), "ABCBC", 0)).toBe(1);
		expect(findSegment(run("bc"), "ABCBC", 2)).toBe(3);
		expect(findSegment(run("bd"), "ABCBC", 0)).toBe(-1);
	});

	it("finds a run with a ? with Shift-And, leftmost from the start position", () => {
		expect(findSegment(run("b?b"), "ABCBC", 0)).toBe(1);
		expect(findSegment(run("?c"), "ABCBC", 3)).toBe(3);
		expect(findSegment(run("b?d"), "ABCBC", 0)).toBe(-1);
		expect(findSegment(run("??"), "A", 0)).toBe(-1);
		expect(findSegment(run("?"), "AB", 2)).toBe(-1);
	});

	it("finds a long run with no ? with Shift-And, also where indexOf would be slow", () => {
		const text = `${"A".repeat(300)}B${"A".repeat(300)}`;
		expect(findSegment(run(`b${"a".repeat(252)}`), text, 0)).toBe(300);
		expect(findSegment(run(`ab${"a".repeat(251)}`), text, 0)).toBe(299);
		expect(findSegment(run(`ab${"a".repeat(251)}`), "A".repeat(600), 0)).toBe(-1);
	});

	it("carries the state across 32-bit words for long runs", () => {
		for (const length of [31, 32, 33, 63, 64, 65, 255]) {
			const pattern = `${"?".repeat(length - 1)}b`;
			const text = `${"a".repeat(300)}b${"a".repeat(10)}`;
			expect(findSegment(run(pattern), text.toUpperCase(), 0)).toBe(301 - length);
			expect(findSegment(run(pattern), "A".repeat(400), 0)).toBe(-1);
		}
	});
});

describe("globMatch", () => {
	it("reads a hand-built Glob with no runs as the empty pattern", () => {
		expect(globMatch("", { segments: [], never: false })).toBe(true);
		expect(globMatch("a", { segments: [], never: false })).toBe(false);
	});

	it("anchors the first run at the start and the last at the end", () => {
		expect(match("abc", "a*")).toBe(true);
		expect(match("abc", "b*")).toBe(false);
		expect(match("abc", "*c")).toBe(true);
		expect(match("abc", "*b")).toBe(false);
		expect(match("abc", "a?c")).toBe(true);
		expect(match("abcd", "a?c")).toBe(false);
	});

	it("finds each middle run after the one before, never overlapping the last", () => {
		expect(match("xaybz", "x*a*b*z")).toBe(true);
		expect(match("xbz", "x*a*z")).toBe(false);
		expect(match("xaz", "x*az*z")).toBe(false);
		expect(match("ab", "a*b")).toBe(true);
		expect(match("a", "a*a")).toBe(false);
		expect(match("aa", "a*a")).toBe(true);
	});

	it("ignores case, and reads ~ as an escape", () => {
		expect(match("ABC", "a*c")).toBe(true);
		expect(match("a*c", "a~*c")).toBe(true);
		expect(match("abc", "a~*c")).toBe(false);
		expect(match("a?", "?~?")).toBe(true);
		expect(match("a~b", "a~b")).toBe(true);
	});

	it("never lets ss match *ß*, and keeps the old miss where ß does not match *ß*", () => {
		expect(match("ss", "*ß*")).toBe(false);
		expect(match("SS", "*ß*")).toBe(false);
		expect(match("ß", "*ß*")).toBe(false);
		expect(match("ß", "S?")).toBe(true);
		expect(match("ﬁ", "f?")).toBe(true);
	});

	it("gives the result the matcher before F11.68 gave, on random patterns and texts", () => {
		// A seeded generator (MINSTD) keeps the case set fixed.
		let seed = 11;
		const next = (limit: number) => {
			seed = (seed * 48271) % 2147483647;
			return seed % limit;
		};
		const alphabet = ["a", "b", "A", "ß", "s", "S", "ﬁ", "f", "i", "~", "𐐨"];
		const symbols = [...alphabet, "*", "?", "*", "?"];
		const pick = (from: readonly string[], length: number) =>
			Array.from({ length }, () => from[next(from.length)] ?? "").join("");
		let matched = 0;
		for (let k = 0; k < 20_000; k++) {
			const text = pick(alphabet, next(12));
			const pattern = pick(symbols, next(9));
			const expected = referenceMatch(text, pattern);
			expect([text, pattern, match(text, pattern)]).toEqual([text, pattern, expected]);
			if (expected) matched++;
		}
		// Long cases: runs over 32 characters, so the Shift-And state spans words.
		for (let k = 0; k < 300; k++) {
			const text = pick(["a", "b"], 200 + next(200));
			const pattern = `*${pick(["a", "b", "?", "?"], 30 + next(60))}*${pick(["a", "?"], next(40))}`;
			expect(match(text, pattern)).toBe(referenceMatch(text, pattern));
		}
		// Both outcomes were checked.
		expect(matched).toBeGreaterThan(500);
	});

	it("costs time in proportion to the text within the cap, for any 255-character pattern", () => {
		for (const pattern of [
			`*${"a?".repeat(126)}b*`,
			`*${"a".repeat(252)}b*`,
			`*${"?".repeat(252)}b*`,
			`*ab${"a".repeat(251)}*`,
		]) {
			const glob = compileGlob(pattern);
			const ratio = bestDoublingRatio((n) => {
				const text = "a".repeat(n);
				return () => globMatch(text, glob);
			}, 16_000);
			expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
		}
	}, 120_000);
});
