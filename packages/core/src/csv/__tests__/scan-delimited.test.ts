import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { retainedHeapBytes } from "../../ooxml/__tests__/retained-heap";
import { parseDelimited, scanDelimited, sniffDelimiter } from "../parse";

// F11.58: the scanner walks the text by index and slices each field, and the sniffer counts the first
// line in place. These cases pin the edges that rewrite touches. Each matches the previous
// character-by-character scanner, and Python's csv.reader (newline='') unless a test says otherwise.

describe("scanDelimited", () => {
	it("reports each field with its 0-based row and column, in order", () => {
		const seen: [string, number, number][] = [];
		scanDelimited('a,"b""c"\n1', ",", (field, row, col) => seen.push([field, row, col]));
		expect(seen).toEqual([
			["a", 0, 0],
			['b"c', 0, 1],
			["1", 1, 0],
		]);
	});

	it("ends with one empty field after a trailing delimiter, and none after a trailing newline", () => {
		const seen: [string, number, number][] = [];
		scanDelimited("a,\r\nb,", ",", (field, row, col) => seen.push([field, row, col]));
		expect(seen).toEqual([
			["a", 0, 0],
			["", 0, 1],
			["b", 1, 0],
			["", 1, 1],
		]);
	});

	it("reports nothing for empty or BOM-only input", () => {
		const seen: string[] = [];
		scanDelimited("", ",", (field) => seen.push(field));
		scanDelimited("\ufeff", ",", (field) => seen.push(field));
		expect(seen).toEqual([]);
	});
});

describe("parseDelimited — quoted-field edges, as Python's csv reads them", () => {
	const cases: [name: string, input: string, expected: string[][]][] = [
		["only an escaped quote", '""""', [['"']]],
		["several escaped quotes", '"a""b""c",d', [['a"b"c', "d"]]],
		["escaped quotes at both ends", '"""a"""', [['"a"']]],
		["an escaped quote ending a field", 'a,"b""",c', [["a", 'b"', "c"]]],
		["an escaped quote before the next quoted field", '"a""","b"', [['a"', "b"]]],
		["two quoted empty fields", '"",""', [["", ""]]],
		["an unterminated quote keeps its escaped quotes", '"a""b', [['a"b']]],
		["quotes after a closing quote are literal", '"a"b"c",d', [['ab"c"', "d"]]],
		["text after a quoted empty field", '""a,b', [["a", "b"]]],
		["CRLF inside a quoted field", '"x\r\ny",z', [["x\r\ny", "z"]]],
		["a quoted newline alone", 'x,"\n",y', [["x", "\n", "y"]]],
		["quoted fields across CRLF rows", '"a,b"\r\n"c""d",e\r\n', [["a,b"], ['c"d', "e"]]],
	];
	for (const [name, input, expected] of cases) {
		it(name, () => {
			expect(parseDelimited(input, ",")).toEqual(expected);
		});
	}

	it("splits on another single-character delimiter", () => {
		expect(parseDelimited('a|"b|c"|d', "|")).toEqual([["a", "b|c", "d"]]);
	});
});

describe("parseDelimited — openjsxl's own rules", () => {
	it("reads a blank line as one empty field (Python yields an empty row)", () => {
		expect(parseDelimited("a\r\r\nb", ",")).toEqual([["a"], [""], ["b"]]);
	});

	it("never splits on a longer, an empty or a non-string delimiter", () => {
		expect(parseDelimited("a,b", ",,")).toEqual([["a,b"]]);
		expect(parseDelimited("a,b", "")).toEqual([["a,b"]]);
		// A plain-JS caller can pass anything; the character-by-character scanner never matched it.
		expect(parseDelimited("a,b", null as never)).toEqual([["a,b"]]);
		expect(parseDelimited("a,b", new String(",") as never)).toEqual([["a,b"]]);
		expect(parseDelimited("a,b", [","] as never)).toEqual([["a,b"]]);
	});
});

describe("sniffDelimiter — the first line, counted in place", () => {
	it("ignores every later line", () => {
		expect(sniffDelimiter("a,b\nc;d;e;f")).toBe(",");
		expect(sniffDelimiter("a,b\r\nc;d;e;f")).toBe(",");
	});

	it("drops the candidates inside a closed quote, even across a line end", () => {
		expect(sniffDelimiter('"a;b\nc;d",e')).toBe(",");
	});

	it("still counts the candidates after an unmatched last quote", () => {
		expect(sniffDelimiter('x,y"a;b;c')).toBe(";");
		expect(sniffDelimiter('x,y"a;b\nc;d')).toBe(";");
	});

	it("breaks ties toward comma, then tab", () => {
		expect(sniffDelimiter("a,b;c")).toBe(",");
		expect(sniffDelimiter("a\tb;c")).toBe("\t");
	});
});

describe("scanner cost", () => {
	// A field built a character at a time is a rope of one node per character, about 32 bytes each.
	it("keeps about the input's size for long fields", () => {
		const row = `${"lorem ipsum dolor sit amet ".repeat(4)},"${"quoted, text ".repeat(8)}",x\n`;
		const text = Array.from({ length: 10_000 }, () => row).join("");
		const retained = retainedHeapBytes(() => parseDelimited(text, ","));
		expect(retained).toBeLessThan(2 * text.length);
	});

	it("keeps about the input's size for quoted fields full of escaped quotes", () => {
		const text = `"${Array.from({ length: 500_000 }, () => 'ab""').join("")}",x\n`;
		const retained = retainedHeapBytes(() => parseDelimited(text, ","));
		expect(retained).toBeLessThan(2 * text.length);
	});

	it("parses in time linear in the input", () => {
		const row = 'customer_1,"Street 9, Apt 3","say ""hi"" to all",note_xyz,42\n';
		const ratio = bestDoublingRatio((n) => {
			const text = row.repeat(n);
			return () => parseDelimited(text, ",");
		}, 50_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 60_000);
});
