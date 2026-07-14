import { describe, expect, it } from "vitest";
import { decodeXmlEntities } from "../entities";

describe("decodeXmlEntities", () => {
	it("returns the input unchanged when there is no ampersand", () => {
		expect(decodeXmlEntities("plain text")).toBe("plain text");
	});

	it("decodes the five predefined entities", () => {
		expect(decodeXmlEntities("a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;")).toBe(
			"a & b < c > d \"e\" 'f'",
		);
	});

	it("decodes decimal character references", () => {
		expect(decodeXmlEntities("line&#10;break")).toBe("line\nbreak");
		expect(decodeXmlEntities("&#65;&#66;&#67;")).toBe("ABC");
	});

	it("decodes hexadecimal character references, including astral code points", () => {
		expect(decodeXmlEntities("&#x41;&#x42;")).toBe("AB");
		expect(decodeXmlEntities("&#x1F600;")).toBe("😀");
		// decimal and hex of the same code point agree
		expect(decodeXmlEntities("&#128512;")).toBe(decodeXmlEntities("&#x1F600;"));
	});

	it("handles adjacent entities", () => {
		expect(decodeXmlEntities("&lt;&gt;&amp;")).toBe("<>&");
	});

	it("leaves surrogate and out-of-range numeric refs literal (no ill-formed scalars)", () => {
		expect(decodeXmlEntities("&#xD800;")).toBe("&#xD800;"); // lone high surrogate
		expect(decodeXmlEntities("&#xDFFF;")).toBe("&#xDFFF;"); // lone low surrogate
		expect(decodeXmlEntities("&#x110000;")).toBe("&#x110000;"); // > U+10FFFF
	});

	it("leaves unknown or malformed entities intact", () => {
		expect(decodeXmlEntities("&unknown;")).toBe("&unknown;");
		expect(decodeXmlEntities("&#xZZ;")).toBe("&#xZZ;"); // not hex digits
		expect(decodeXmlEntities("5 & 6")).toBe("5 & 6"); // bare ampersand
		expect(decodeXmlEntities("a &amp b")).toBe("a &amp b"); // missing semicolon
	});

	it("leaves an over-long numeric reference literal — bounded to fit the stream window (F9.7)", () => {
		// The digit count is capped at 6 hex / 7 decimal (enough for U+10FFFF); a longer run, even
		// pure leading-zero padding, stays literal so a straddling reference decodes the same in the
		// streaming and one-shot tokenizers (they must agree on a maximum entity length).
		expect(decodeXmlEntities("&#x10FFFF;")).toBe("\u{10FFFF}"); // the longest DECODABLE hex ref
		expect(decodeXmlEntities("&#1114111;")).toBe("\u{10FFFF}"); // the longest DECODABLE decimal ref
		expect(decodeXmlEntities("&#x0000041;")).toBe("&#x0000041;"); // 7 hex digits → literal
		expect(decodeXmlEntities("&#00000065;")).toBe("&#00000065;"); // 8 decimal digits → literal
		expect(decodeXmlEntities("&#x0000000000000041;")).toBe("&#x0000000000000041;");
	});
});
