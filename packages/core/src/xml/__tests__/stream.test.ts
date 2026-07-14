import { describe, expect, it } from "vitest";
import { createXmlStream } from "../stream";
import { tokenize, type XmlToken } from "../tokenizer";

// Streaming splits a text run into several `text` tokens at chunk boundaries; merging
// adjacent text tokens makes the streamed output directly comparable to the one-shot output.
function mergeText(tokens: XmlToken[]): XmlToken[] {
	const out: XmlToken[] = [];
	for (const token of tokens) {
		const prev = out[out.length - 1];
		if (token.kind === "text" && prev?.kind === "text") {
			out[out.length - 1] = { kind: "text", value: prev.value + token.value };
		} else {
			out.push(token);
		}
	}
	// A merge can leave an empty text token only if the input had one; drop empties for a
	// stable comparison (one-shot tokenize never emits empty text runs).
	return out.filter((t) => !(t.kind === "text" && t.value === ""));
}

function streamInChunks(xml: string, size: number): XmlToken[] {
	const stream = createXmlStream();
	const tokens: XmlToken[] = [];
	for (let i = 0; i < xml.length; i += size) {
		tokens.push(...stream.push(xml.slice(i, i + size)));
	}
	tokens.push(...stream.flush());
	return tokens;
}

describe("createXmlStream", () => {
	const samples = [
		'<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
		'<c r="A1"><v>a &amp; b &lt;c&gt;</v></c>',
		"<is><t>Hello &amp; welcome</t><![CDATA[raw <> data]]></is>",
		'<!-- a comment with < and > inside --><row r="2"><c r="A2"/></row>',
		"<a x=\"quoted > gt\" y='apos > gt'>text</a>",
		'<?xml version="1.0"?><root>\n\t<leaf/>\n</root>',
	];

	for (const xml of samples) {
		it(`matches one-shot tokenize for every chunk size: ${xml.slice(0, 32)}…`, () => {
			const expected = mergeText([...tokenize(xml)]);
			for (let size = 1; size <= xml.length; size++) {
				expect(mergeText(streamInChunks(xml, size)), `chunk size ${size}`).toEqual(
					expected,
				);
			}
		});
	}

	it("keeps an entity split across chunks intact", () => {
		const stream = createXmlStream();
		const tokens = [...stream.push("<v>a &am"), ...stream.push("p; b</v>"), ...stream.flush()];
		expect(mergeText(tokens)).toEqual([
			{ kind: "open", name: "v", attrs: {}, selfClosing: false },
			{ kind: "text", value: "a & b" },
			{ kind: "close", name: "v" },
		]);
	});

	it("keeps a tag split across chunks intact", () => {
		const stream = createXmlStream();
		const tokens = [...stream.push('<c r="A'), ...stream.push('1" t="s"/>'), ...stream.flush()];
		expect(tokens).toEqual([
			{ kind: "open", name: "c", attrs: { r: "A1", t: "s" }, selfClosing: true },
		]);
	});

	it("emits a stray literal & in a long run without stalling (bounded buffer)", () => {
		// A "&" that no ";" ever completes must not pin the buffer; the merged text must still
		// match one-shot tokenize across every chunk size.
		const xml = `<v>${"x".repeat(40)} & ${"y".repeat(40)}</v>`;
		const expected = mergeText([...tokenize(xml)]);
		for (let size = 1; size <= 8; size++) {
			expect(mergeText(streamInChunks(xml, size)), `chunk size ${size}`).toEqual(expected);
		}
	});

	it("buffers until a construct completes (no partial emission)", () => {
		const stream = createXmlStream();
		expect(stream.push("<!-- unterminated")).toEqual([]);
		expect(stream.push(" still going")).toEqual([]);
		expect(stream.push(" -->")).toEqual([]);
		expect(stream.push("<x/>")).toEqual([
			{ kind: "open", name: "x", attrs: {}, selfClosing: true },
		]);
	});
});

// F9.7 — the resumable scan. A construct straddling many chunks is searched ONCE across all
// pushes (a carried cursor), not rescanned from its start each push (the old O(n²) hostile-input
// blowup), and a single construct can't pin the buffer without bound.
describe("createXmlStream — F9.7 resumable scan", () => {
	it("finds a comment/CDATA terminator split across the resume boundary", () => {
		// "--" arrives in one push, ">" in the next — the resume must back up to see "-->".
		const stream = createXmlStream();
		expect(stream.push("<!--body--")).toEqual([]);
		expect(stream.push("><y/>")).toEqual([
			{ kind: "open", name: "y", attrs: {}, selfClosing: true },
		]);
		const s2 = createXmlStream();
		expect(s2.push("<![CDATA[data]]")).toEqual([]);
		expect(mergeText([...s2.push(">"), ...s2.flush()])).toEqual([
			{ kind: "text", value: "data" },
		]);
	});

	it("carries quote state across pushes — a '>' inside a split attribute value can't end the tag", () => {
		const stream = createXmlStream();
		expect(stream.push('<c v="one > ')).toEqual([]);
		expect(stream.push('two"')).toEqual([]); // quote closes; tag still open
		expect(stream.push("/>")).toEqual([
			{ kind: "open", name: "c", attrs: { v: "one > two" }, selfClosing: true },
		]);
	});

	it("streams a construct split across MANY chunks in linear total work", () => {
		// 4000 × 512B = a 2 MB comment body. The old rescan-from-zero cost ~2 GB of char touches
		// (visible hang, caught by the suite timeout); the carried cursor makes it one 2 MB scan.
		const stream = createXmlStream();
		expect(stream.push("<!--")).toEqual([]);
		const filler = "c".repeat(512);
		for (let n = 0; n < 4000; n++) {
			expect(stream.push(filler)).toEqual([]);
		}
		expect(stream.push("--><done/>")).toEqual([
			{ kind: "open", name: "done", attrs: {}, selfClosing: true },
		]);
	});

	it("streams a giant TAG (attribute value) across many chunks the same way", () => {
		const stream = createXmlStream();
		expect(stream.push('<c r="A1" v="')).toEqual([]);
		const filler = "a".repeat(512);
		for (let n = 0; n < 2000; n++) {
			expect(stream.push(filler)).toEqual([]);
		}
		const tokens = stream.push('"/>');
		expect(tokens).toHaveLength(1);
		expect(tokens[0]?.kind).toBe("open");
	});

	it("streams a giant '<!…>' declaration in LINEAR time (F9.7 review: the 1-char-terminator O(n²) bug)", () => {
		// The `<!DOCTYPE …>` (and `</…>`) terminator is a single '>', so its resume overlap must be
		// EMPTY — a `slice(-0)` bug kept the whole window and re-searched it every push (quadratic).
		// A 32 MiB unterminated declaration must stream well under a second (it was ~5.5s when quadratic).
		const stream = createXmlStream();
		stream.push("<!DOCTYPE ");
		const filler = "x".repeat(64 * 1024);
		const t0 = performance.now();
		for (let n = 0; n < 512; n++) expect(stream.push(filler)).toEqual([]); // 32 MiB, no '>'
		expect(performance.now() - t0).toBeLessThan(1500);
		expect(stream.push(">ok<x/>").at(-1)?.kind).toBe("open");
	});

	it("fails typed once a single unterminated construct exceeds MAX_UNFINISHED_CONSTRUCT", () => {
		const stream = createXmlStream();
		stream.push("<!--");
		// Feed 8 MiB chunks of comment body; the cap must trip before ~72 MiB, typed.
		const chunk = "z".repeat(8 * 1024 * 1024);
		expect(() => {
			for (let n = 0; n < 10; n++) stream.push(chunk);
		}).toThrowError(
			expect.objectContaining({ name: "XlsxError", code: "part-too-large" }) as Error,
		);
	});

	it("drained text never counts toward the construct cap (a huge sheet still streams)", () => {
		// Long runs of complete constructs + text keep the buffer near-empty regardless of total
		// volume — only a single UNFINISHED construct can accumulate.
		const stream = createXmlStream();
		const row = `<row><c t="inlineStr"><is><t>${"v".repeat(1024)}</t></is></c></row>`;
		for (let n = 0; n < 2000; n++) {
			expect(stream.push(row).length).toBeGreaterThan(0); // tokens flow, buffer drains
		}
		expect(stream.flush()).toEqual([]);
	});

	// The rewrite's boundary detection (safeBoundary/scanTag) must agree with the one-shot
	// tokenizer even on MALFORMED markup — a `<` that is literal text, a `'`/`>` inside a tag name,
	// a `<!DOCTYPE>` declaration — or a split there would leak/absorb content one-shot would not.
	// These are the exact shapes a differential fuzz surfaced (all now match at every chunk size).
	it("matches one-shot on malformed markup the boundary heuristic used to split wrong", () => {
		const samples = [
			"a<=b", // '<' followed by '=' is literal text, not a tag
			"x < y", // '<' followed by space is literal text
			"<<b='>'/>", // a '>' and quote inside the tag NAME — name ends at the first '>'
			"<a b='>' c=\"<\">t</a>", // quotes protect '>' only inside a value (after '=')
			"<!DOCTYPE html>t", // a declaration ends at the first plain '>'
			"<!x='>'>after", // a declaration's '>' is quote-UNAWARE
			"<b='>'<!--]]>  <?x", // the fuzz's minimised regressions
			">--\n<\n<!--]]></a>",
			'& <="x"<?<?]]>\t',
			"text<?pi never closes",
			'</a b="x><t q=">"/>', // F9.7 REVIEW: a close tag ends at the first plain '>' (quote-UNAWARE)
			"</a '>' b>c", // a quote inside a close-tag name doesn't protect '>'
			"t&#x0000000000000041;u", // F9.7 REVIEW: an over-long numeric charref stays literal in both paths
			"t&#00000000000000065;u",
			"ok&#x41;&#160;done", // valid charrefs still decode when split anywhere
			"<!DOCTYPE x '>' y>z", // a declaration '>' is quote-unaware
		];
		for (const xml of samples) {
			const expected = mergeText([...tokenize(xml)]);
			for (let size = 1; size <= xml.length; size++) {
				expect(
					mergeText(streamInChunks(xml, size)),
					`${JSON.stringify(xml)} @ ${size}`,
				).toEqual(expected);
			}
		}
	});

	it("matches one-shot tokenize over a deterministic malformed-markup fuzz", () => {
		// xorshift32 — deterministic, no Math.random. Builds strings from fragments dense in the
		// tricky shapes (unclosed constructs, stray terminators, quotes/`>` in names, declarations)
		// and asserts streamed==one-shot at several random chunk splittings.
		let st = 0x51ed2718;
		const rand = () => {
			st ^= st << 13;
			st ^= st >>> 17;
			st ^= st << 5;
			return (st >>> 0) / 0x100000000;
		};
		const frags = [
			"<a",
			"<a>",
			"</a>",
			"<b/>",
			"<!--",
			"-->",
			"<![CDATA[",
			"]]>",
			"<?",
			"?>",
			'="x"',
			"='y'",
			">",
			"<",
			"&amp;",
			"& ",
			"text ",
			"]]",
			"--",
			"<!",
			"<![",
			"<!DOCTYPE>",
			"a",
			"b='>'",
			'c="<"',
			"  ",
			"\n",
			'</a b=">"',
			"</x '>'",
			"&#x0000041;",
			"&#00000065;",
			"&#x41;",
		];
		for (let iter = 0; iter < 3000; iter++) {
			let xml = "";
			const parts = 1 + Math.floor(rand() * 14);
			for (let p = 0; p < parts; p++) xml += frags[Math.floor(rand() * frags.length)];
			const expected = mergeText([...tokenize(xml)]);
			for (let s = 0; s < 3; s++) {
				const size = 1 + Math.floor(rand() * 5);
				expect(mergeText(streamInChunks(xml, size)), JSON.stringify(xml)).toEqual(expected);
			}
		}
	});
});
