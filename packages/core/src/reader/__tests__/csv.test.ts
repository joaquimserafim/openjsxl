import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { inferCsvValue, parseDelimited, sniffDelimiter } from "../../csv";
import { workbookToInput, writeXlsx } from "../../writer";
import { openCsv } from "../csv";
import { openXlsx } from "../workbook";

// F7.3 — the .csv / .tsv reader. The scanner's splitting/quoting/newline behavior is pinned against
// Python's stdlib `csv` (the RFC 4180 reference — ground truth captured with newline=''); type
// inference and delimiter sniffing are openjsxl's own documented rules.

describe("parseDelimited — matches Python's stdlib csv", () => {
	// Each expected value is exactly `list(csv.reader(io.StringIO(input, newline='')))`.
	const cases: [name: string, input: string, expected: string[][]][] = [
		["simple", "a,b,c", [["a", "b", "c"]]],
		[
			"two rows",
			"a,b,c\n1,2,3",
			[
				["a", "b", "c"],
				["1", "2", "3"],
			],
		],
		["quoted delimiter", '"a,b",c', [["a,b", "c"]]],
		["escaped quote", '"a""b",c', [['a"b', "c"]]],
		["embedded newline", '"line1\nline2",b', [["line1\nline2", "b"]]],
		[
			"CRLF + trailing",
			"a,b\r\n1,2\r\n",
			[
				["a", "b"],
				["1", "2"],
			],
		],
		[
			"lone CR (old Mac)",
			"a,b\rc,d",
			[
				["a", "b"],
				["c", "d"],
			],
		],
		["empty middle field", "a,,c", [["a", "", "c"]]],
		["trailing delimiter", "a,b,", [["a", "b", ""]]],
		["trailing newline → no empty row", "a,b\n", [["a", "b"]]],
		[
			"ragged rows",
			"a,b,c\n1,2",
			[
				["a", "b", "c"],
				["1", "2"],
			],
		],
		["quoted empty field", '""', [[""]]],
		["unterminated quote → one field to EOF", '"open,b\nc', [["open,b\nc"]]],
		// A quote is special only at field start; mid-field quotes are literal (RFC 4180 / Python csv).
		["mid-field quote is literal", 'a"b,c', [['a"b', "c"]]],
		["quote then trailing text", '"a"b,c', [["ab", "c"]]],
		["space before quote → literal", ' "x",y', [[' "x"', "y"]]],
		["quotes around mid-field text", 'ab"cd"ef,g', [['ab"cd"ef', "g"]]],
	];
	for (const [name, input, expected] of cases) {
		it(name, () => {
			expect(parseDelimited(input, ",")).toEqual(expected);
		});
	}

	it("strips a leading BOM", () => {
		expect(parseDelimited("﻿a,b", ",")).toEqual([["a", "b"]]);
	});

	it("empty input yields no rows", () => {
		expect(parseDelimited("", ",")).toEqual([]);
	});
});

describe("inferCsvValue — conservative type inference (never dates)", () => {
	const num = (v: number) => ({ type: "number", value: v });
	it("infers plain numbers", () => {
		expect(inferCsvValue("42")).toEqual(num(42));
		expect(inferCsvValue("-3.5")).toEqual(num(-3.5));
		expect(inferCsvValue("+7")).toEqual(num(7));
		expect(inferCsvValue("0")).toEqual(num(0));
		expect(inferCsvValue("0.25")).toEqual(num(0.25));
		expect(inferCsvValue(".5")).toEqual(num(0.5));
		expect(inferCsvValue("1e3")).toEqual(num(1000));
	});
	it("infers booleans (case-insensitive)", () => {
		expect(inferCsvValue("TRUE")).toEqual({ type: "boolean", value: true });
		expect(inferCsvValue("false")).toEqual({ type: "boolean", value: false });
		expect(inferCsvValue("True")).toEqual({ type: "boolean", value: true });
	});
	it("keeps leading-zero digits as strings (ZIP/ID preservation)", () => {
		expect(inferCsvValue("007")).toEqual({ type: "string", value: "007" });
		expect(inferCsvValue("00")).toEqual({ type: "string", value: "00" });
	});
	it("keeps a big integer beyond safe range as a string (no precision loss)", () => {
		expect(inferCsvValue("12345678901234567890")).toEqual({
			type: "string",
			value: "12345678901234567890",
		});
		// 2^53 + 1 is not exactly representable → string; 2^53 - 1 (MAX_SAFE) is fine → number.
		expect(inferCsvValue("9007199254740993")?.type).toBe("string");
		expect(inferCsvValue("9007199254740991")).toEqual(num(9007199254740991));
	});
	it("never infers dates, Infinity/NaN, hex, or thousands separators", () => {
		expect(inferCsvValue("01/02/2024")).toEqual({ type: "string", value: "01/02/2024" });
		expect(inferCsvValue("2024-01-15")).toEqual({ type: "string", value: "2024-01-15" });
		expect(inferCsvValue("NaN")?.type).toBe("string");
		expect(inferCsvValue("Infinity")?.type).toBe("string");
		expect(inferCsvValue("0x1F")?.type).toBe("string");
		expect(inferCsvValue("1,000")?.type).toBe("string");
	});
	it("treats an empty field as an empty cell", () => {
		expect(inferCsvValue("")).toBeUndefined();
	});
});

describe("sniffDelimiter", () => {
	it("picks the most frequent candidate outside quotes", () => {
		expect(sniffDelimiter("a,b,c")).toBe(",");
		expect(sniffDelimiter("a\tb\tc")).toBe("\t");
		expect(sniffDelimiter("a;b;c")).toBe(";");
		expect(sniffDelimiter('"a,b,c";d;e')).toBe(";"); // commas inside quotes don't count
		expect(sniffDelimiter("")).toBe(","); // default
	});
});

describe("openCsv — end to end", () => {
	it("reads a small CSV into typed cells with a synthesized dimension", () => {
		const wb = openCsv("name,qty,active\nApples,42,TRUE\nPears,7,false");
		const s = wb.sheet("Sheet1");
		expect(s.cell("A1")).toEqual({ ref: "A1", type: "string", value: "name" });
		expect(s.cell("B2")).toEqual({ ref: "B2", type: "number", value: 42 });
		expect(s.cell("C2")).toEqual({ ref: "C2", type: "boolean", value: true });
		expect(s.cell("C3")).toEqual({ ref: "C3", type: "boolean", value: false });
		expect(s.dimension).toBe("A1:C3");
	});

	it("honors delimiter, sheetName, and inferTypes options", () => {
		const tsv = openCsv("a\t42\tx", { delimiter: "\t", sheetName: "Data", inferTypes: false });
		const s = tsv.sheet("Data");
		expect(tsv.sheets[0]?.name).toBe("Data");
		expect(s.cell("B1")).toEqual({ ref: "B1", type: "string", value: "42" }); // inference off
	});

	it("accepts raw bytes and strips a BOM", () => {
		const bytes = new TextEncoder().encode("﻿hello,world");
		const s = openCsv(bytes).sheet("Sheet1");
		expect(s.cell("A1").value).toBe("hello");
		expect(s.cell("B1").value).toBe("world");
	});

	it("streams rows sparsely (empty fields are absent cells)", async () => {
		const s = openCsv("a,,c").sheet("Sheet1");
		const rows = [];
		for await (const row of s.rows()) rows.push(row.cells.map((c) => c.ref));
		expect(rows).toEqual([["A1", "C1"]]); // B1 (empty) is absent
		expect(s.cell("B1").type).toBe("empty");
	});

	it("degrades unsupported accessors", async () => {
		const s = openCsv("a,b").sheet("Sheet1");
		expect(s.mergedCells).toEqual([]);
		expect(s.style("A1")).toBeUndefined();
		expect(s.numberFormat("A1")).toBeUndefined();
		expect(s.formula("A1")).toBeUndefined();
		expect(s.hyperlinks).toEqual([]);
		expect(await s.images()).toEqual([]);
		expect(s.freeze).toBeUndefined();
	});

	it("reads an empty input as a one-sheet workbook with no cells", () => {
		const wb = openCsv("");
		expect(wb.sheets.map((x) => x.name)).toEqual(["Sheet1"]);
		expect(wb.sheet("Sheet1").dimension).toBeUndefined();
		expect(wb.sheet("Sheet1").cell("A1").type).toBe("empty");
	});

	it("converts a CSV to .xlsx through the bridge, values intact", async () => {
		const wb = openCsv("name,qty\nApples,42");
		const out = await openXlsx(await writeXlsx(await workbookToInput(wb)));
		const s = out.sheet("Sheet1");
		expect(s.cell("A1").value).toBe("name");
		expect(s.cell("B2").value).toBe(42);
	});

	it("reads the committed basic.csv fixture (matches Python csv + inference)", async () => {
		const s = openCsv(await loadFixture("basic.csv")).sheet("Sheet1");
		expect(s.cell("A2").value).toBe("007"); // leading zero preserved
		expect(s.cell("B2").value).toBe("Acme, Inc."); // quoted delimiter
		expect(s.cell("C2").value).toBe(42); // number
		expect(s.cell("D2").value).toBe(true); // boolean
		expect(s.cell("B3").value).toBe("multi\nline"); // embedded newline in a quoted field
		expect(s.cell("B4").value).toBe('a "quoted" word'); // escaped quote
		expect(s.dimension).toBe("A1:D4");
	});
});
