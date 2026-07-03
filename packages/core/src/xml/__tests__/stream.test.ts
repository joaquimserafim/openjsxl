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
