import { describe, expect, it } from "vitest";
import {
	escapeAttr,
	escapeText,
	isPlainRecord,
	needsPreserve,
	ownValue,
	preserveAttr,
} from "../xml";

describe("escapeText", () => {
	it("escapes the markup-significant characters", () => {
		expect(escapeText("a & b < c > d")).toBe("a &amp; b &lt; c &gt; d");
	});

	it("escapes & first so introduced entities are not double-escaped", () => {
		expect(escapeText("<&>")).toBe("&lt;&amp;&gt;");
	});

	it("preserves CR as a character reference without changing CRLF's LF or a lone LF", () => {
		expect(escapeText("a\rb\r\nc\nd")).toBe("a&#13;b&#13;\nc\nd");
		expect(escapeText("literal &#13; then \r")).toBe("literal &amp;#13; then &#13;");
	});

	it("leaves quotes untouched (text context)", () => {
		expect(escapeText('say "hi"')).toBe('say "hi"');
	});
});

describe("escapeAttr", () => {
	it("also escapes double quotes", () => {
		expect(escapeAttr('a "b" & <c>')).toBe("a &quot;b&quot; &amp; &lt;c&gt;");
	});

	it("preserves attribute whitespace with character references", () => {
		expect(escapeAttr("a\tb\nc\rd")).toBe("a&#9;b&#10;c&#13;d");
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

describe("isPlainRecord", () => {
	it("accepts a direct Object instance, including a null-prototype object", () => {
		expect(isPlainRecord({})).toBe(true);
		expect(isPlainRecord({ a: 1 })).toBe(true);
		// Object.create(null) has no prototype — a legitimate bag of properties the validators must
		// accept (it is the safest shape a caller can pass; rejecting it would be a false negative).
		expect(isPlainRecord(Object.create(null))).toBe(true);
	});

	it("rejects non-objects and null", () => {
		expect(isPlainRecord(null)).toBe(false);
		expect(isPlainRecord(undefined)).toBe(false);
		expect(isPlainRecord(42)).toBe(false);
		expect(isPlainRecord("str")).toBe(false);
		expect(isPlainRecord(true)).toBe(false);
		expect(isPlainRecord(Symbol("s"))).toBe(false);
		expect(isPlainRecord(() => {})).toBe(false);
	});

	it("rejects arrays and exotic objects — the prototype-smuggling defense", () => {
		// An array's prototype is Array.prototype, not Object.prototype — writer input that expects a
		// record must not silently accept a list. Dates/Maps/class instances likewise carry methods and
		// getters that could smuggle values past single-read validation, so they are rejected up front.
		expect(isPlainRecord([])).toBe(false);
		expect(isPlainRecord([1, 2])).toBe(false);
		expect(isPlainRecord(new Date())).toBe(false);
		expect(isPlainRecord(new Map())).toBe(false);
		expect(isPlainRecord(/re/)).toBe(false);
		class Widget {
			x = 1;
		}
		expect(isPlainRecord(new Widget())).toBe(false);
	});

	it("rejects an object whose prototype was reassigned to a non-Object prototype", () => {
		const spoofed = Object.setPrototypeOf({ a: 1 }, Array.prototype);
		expect(isPlainRecord(spoofed)).toBe(false);
	});
});

describe("ownValue", () => {
	it("retains own undefined presence, non-enumerable values, and null prototypes", () => {
		const record: Record<string, unknown> = { present: undefined };
		Object.setPrototypeOf(record, null);
		Object.defineProperty(record, "hidden", { value: 7 });
		expect(ownValue(record, "missing")).toBeUndefined();
		expect(ownValue(record, "present")).toBeUndefined();
		expect(Object.hasOwn(record, "missing")).toBe(false);
		expect(Object.hasOwn(record, "present")).toBe(true);
		expect(ownValue(record, "hidden")).toBe(7);
	});

	it("reads an own getter once with its original receiver", () => {
		let reads = 0;
		const record = {
			marker: 4,
			get value() {
				reads++;
				return this.marker;
			},
		};
		expect(ownValue(record, "value")).toBe(4);
		expect(reads).toBe(1);
	});

	it("never reads inherited getters, including sparse array entries", () => {
		let reads = 0;
		const record: Record<string, unknown> = {};
		Object.setPrototypeOf(record, {
			get value() {
				reads++;
				return 9;
			},
		});
		expect(ownValue(record, "value")).toBeUndefined();
		const rows: number[] = new Array(1);
		Object.setPrototypeOf(rows, {
			get 0() {
				reads++;
				return 9;
			},
		});
		expect(ownValue(rows, 0)).toBeUndefined();
		expect(reads).toBe(0);
	});
});
