import { describe, expect, it } from "vitest";
import { escapeAttr, escapeText, needsPreserve, preserveAttr } from "../xml";

describe("escapeText", () => {
	it("escapes the markup-significant characters", () => {
		expect(escapeText("a & b < c > d")).toBe("a &amp; b &lt; c &gt; d");
	});

	it("escapes & first so introduced entities are not double-escaped", () => {
		expect(escapeText("<&>")).toBe("&lt;&amp;&gt;");
	});

	it("leaves quotes untouched (text context)", () => {
		expect(escapeText('say "hi"')).toBe('say "hi"');
	});
});

describe("escapeAttr", () => {
	it("also escapes double quotes", () => {
		expect(escapeAttr('a "b" & <c>')).toBe("a &quot;b&quot; &amp; &lt;c&gt;");
	});
});

describe("needsPreserve / preserveAttr", () => {
	it("flags leading/trailing whitespace", () => {
		expect(needsPreserve(" hi")).toBe(true);
		expect(needsPreserve("hi ")).toBe(true);
		expect(needsPreserve("a\tb\n")).toBe(true);
		expect(needsPreserve("hi")).toBe(false);
		expect(needsPreserve("a b")).toBe(false);
	});

	it("emits the attribute only when needed", () => {
		expect(preserveAttr(" hi")).toBe(' xml:space="preserve"');
		expect(preserveAttr("hi")).toBe("");
	});
});
