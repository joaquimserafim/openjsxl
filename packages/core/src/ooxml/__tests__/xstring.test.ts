import { describe, expect, it } from "vitest";
import { isXmlSafe } from "../../utils";
import { decodeXstring, encodeXstring } from "../xstring";

// F9.6 — the ST_Xstring escape codec. The load-bearing properties:
//   1. decode(encode(s)) === s for EVERY string (lossless storage round-trip);
//   2. encode is the IDENTITY on strings with nothing to escape (byte-identity for clean files);
//   3. decode is single-pass and non-reentrant (a decoded `_x005F_` is never re-examined) —
//      matching Excel and openpyxl (`re.sub` semantics), which is what makes our bytes agree.

describe("decodeXstring", () => {
	it("decodes a control-char escape the way Excel does", () => {
		expect(decodeXstring("a_x000B_b")).toBe("a\x0Bb");
		expect(decodeXstring("_x0000_")).toBe("\x00");
		expect(decodeXstring("_x0041_")).toBe("A"); // ANY unit decodes, not just controls
		expect(decodeXstring("_x30C6_")).toBe("テ");
	});

	it("accepts hex in either case, but only a lowercase x", () => {
		expect(decodeXstring("_x000b_")).toBe("\x0B");
		expect(decodeXstring("_X000B_")).toBe("_X000B_"); // uppercase X is NOT an escape
	});

	it("leaves non-matching text untouched (identity fast path)", () => {
		for (const s of ["", "plain", "_x", "_x12_", "_x12345_", "_xZZZZ_", "a_b_c", "__"]) {
			expect(decodeXstring(s)).toBe(s);
		}
	});

	it("is non-reentrant: _x005F_ protects a literal escape shape", () => {
		expect(decodeXstring("_x005F_x0041_")).toBe("_x0041_"); // the protected literal survives
		expect(decodeXstring("_x005F_x005F_")).toBe("_x005F_");
	});

	it("scans non-overlapping, left to right (shared underscores consume once)", () => {
		// The trailing _ of the first escape is CONSUMED — the "x0042_" after it is literal text,
		// exactly what a producer's non-overlapping regex substitution yields.
		expect(decodeXstring("_x0041_x0042_")).toBe("Ax0042_");
	});

	it("rebuilds astral characters stored as a surrogate-pair escape sequence", () => {
		expect(decodeXstring("_xD83D__xDE00_")).toBe("😀");
	});
});

describe("encodeXstring", () => {
	it("is the identity on strings with nothing to escape", () => {
		for (const s of [
			"",
			"plain text",
			"tab\there",
			"line\nbreak",
			"😀 emoji",
			"_x12_",
			"a_b",
		]) {
			expect(encodeXstring(s)).toBe(s);
		}
	});

	it("escapes the characters isXmlSafe rejects — and nothing else below 0x20 except CR", () => {
		expect(encodeXstring("a\x0Bb")).toBe("a_x000B_b");
		expect(encodeXstring("\x00")).toBe("_x0000_");
		expect(encodeXstring("\x1F")).toBe("_x001F_");
		expect(encodeXstring("￾￿")).toBe("_xFFFE__xFFFF_");
		expect(encodeXstring("keep \t\n")).toBe("keep \t\n"); // tab/LF are XML-legal — stay raw
	});

	it("escapes CR — XML parsers normalize a raw CR to LF, so Excel stores it as _x000D_", () => {
		expect(encodeXstring("a\rb")).toBe("a_x000D_b");
		expect(encodeXstring("crlf\r\n")).toBe("crlf_x000D_\n");
		expect(decodeXstring(encodeXstring("a\rb"))).toBe("a\rb");
	});

	it("escapes lone surrogates but passes well-formed pairs through", () => {
		expect(encodeXstring("\uD800")).toBe("_xD800_"); // lone high
		expect(encodeXstring("\uDC00")).toBe("_xDC00_"); // lone low
		expect(encodeXstring("😀")).toBe("😀"); // a real pair is untouched
		expect(encodeXstring("a\uD83Db")).toBe("a_xD83D_b"); // high not followed by low
	});

	it("protects a literal _xHHHH_ so Excel cannot decode data into different characters", () => {
		expect(encodeXstring("_x0041_")).toBe("_x005F_x0041_");
		expect(encodeXstring("id_x0001_tail")).toBe("id_x005F_x0001_tail");
		expect(encodeXstring("_x005F_")).toBe("_x005F_x005F_");
	});

	it("protects a literal _xHHHH whose shape an EMITTED escape would complete (review HIGH)", () => {
		// The input has no full shape — but the escape emitted for the following unit starts
		// with `_`, which would close one in the OUTPUT and make Excel decode the literal.
		expect(encodeXstring("_x0041\x01")).toBe("_x005F_x0041_x0001_");
		expect(decodeXstring(encodeXstring("_x0041\x01"))).toBe("_x0041\x01");
		expect(decodeXstring(encodeXstring("_x005F\x01"))).toBe("_x005F\x01");
		expect(decodeXstring(encodeXstring("abc_xBEEF\x00tail"))).toBe("abc_xBEEF\x00tail");
		expect(decodeXstring(encodeXstring("LOG_x0041\x01END"))).toBe("LOG_x0041\x01END");
		expect(decodeXstring(encodeXstring("_x1aFf\uDC00"))).toBe("_x1aFf\uDC00");
		expect(decodeXstring(encodeXstring("_x0041_xa50A\uD800B_x0041_"))).toBe(
			"_x0041_xa50A\uD800B_x0041_",
		);
		expect(decodeXstring(encodeXstring("_x0041\r"))).toBe("_x0041\r"); // CR completes it too
		// A literal `_xHHHH` followed by something SAFE needs no protection (stays identity).
		expect(encodeXstring("_x0041!")).toBe("_x0041!");
		expect(encodeXstring("_x0041😀")).toBe("_x0041😀");
	});

	it("always produces an XML-safe string", () => {
		for (const s of ["\x00\x01\x02", "𐀀\uD800", "_x0000_", "mixed\x0B_x0041_😀"]) {
			expect(isXmlSafe(encodeXstring(s))).toBe(true);
		}
	});
});

describe("the codec round-trip", () => {
	it("decode(encode(s)) === s for adversarial shapes", () => {
		const cases = [
			"",
			"plain",
			"_x0041_",
			"_x005F_",
			"_x005F_x0041_",
			"_x0041_x0042_",
			"__x0041_",
			"_x0041",
			"x0041_",
			"\x00\x0B\x1F",
			"\r",
			"\r\n\r",
			"\uD800",
			"\uDFFF",
			"😀_x0000_😀",
			"_x0041\x01",
			"_x005F\x01",
			"_x4141\x01",
			"_xBEEF\uD800",
			"_x0041\r",
			"_x0041_x0042\x00",
			"a".repeat(1000) + "_x005F_".repeat(50),
			"_".repeat(20),
			"_x".repeat(20),
		];
		for (const s of cases) {
			expect(decodeXstring(encodeXstring(s))).toBe(s);
		}
	});

	it("decode(encode(s)) === s over a deterministic pseudo-random corpus", () => {
		// xorshift32 — deterministic, no Math.random in tests. Skews toward the interesting
		// alphabet: underscores, hex digits, x, controls, surrogate halves.
		let state = 0x9e3779b9;
		const rand = () => {
			state ^= state << 13;
			state ^= state >>> 17;
			state ^= state << 5;
			return (state >>> 0) / 0x100000000;
		};
		const alphabet = [
			"_",
			"x",
			"0",
			"4",
			"F",
			"f",
			"1",
			"A",
			"\x0B",
			"\x00",
			"\r",
			"\uD83D",
			"\uDE00",
			"￾",
			"a",
			"€",
			// Composite fragments so partial escape shapes (the review-HIGH class: a literal
			// `_xHHHH` completed by an emitted escape) appear at high density.
			"_x0041",
			"_x005F",
			"_xBEEF",
			"_x0041_",
		];
		for (let n = 0; n < 500; n++) {
			let s = "";
			const len = Math.floor(rand() * 24);
			for (let k = 0; k < len; k++) {
				const pick = alphabet[Math.floor(rand() * alphabet.length)];
				if (pick !== undefined) s += pick;
			}
			const encoded = encodeXstring(s);
			expect(isXmlSafe(encoded)).toBe(true);
			expect(decodeXstring(encoded)).toBe(s);
		}
	});
});
