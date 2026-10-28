import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.66: text that `&` or a built-in builds past Excel's 32,767-character cell limit is #VALUE!, as in
// Excel, where it used to throw a bare RangeError or return the longer text.

const VALUE = errorValue("#VALUE!");
const lengthOf = (values: readonly unknown[]) =>
	values.map((value) => (typeof value === "string" ? value.length : value));

describe("text past 32,767 characters", () => {
	it("is #VALUE! where building it used to throw", async () => {
		// Each REPT cell holds 32,767 characters; 16,400 of them overflow a JavaScript string.
		const rows = Array.from({ length: 16400 }, () => [
			{ formula: 'REPT("a",32767)', value: 0 },
		]);
		expect(
			await evaluateFormulas(
				[
					'SUBSTITUTE(REPT("a",32767),"a",REPT("b",32767))',
					'TEXTJOIN(",",TRUE,Data!A1:A16400)',
					"CONCAT(Data!A1:A16400)",
				],
				rows,
			),
		).toEqual([VALUE, VALUE, VALUE]);
	});

	it("is #VALUE! at the end of a doubling & chain, which used to throw", async () => {
		const rows = [[{ formula: 'REPT("a",32767)', value: 0 }]];
		for (let i = 1; i < 16; i++) rows.push([{ formula: `A${i}&A${i}`, value: 0 }]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "Data", rows }] }));
		expect(await evaluateCell(book, "Data", "A16")).toEqual(VALUE);
		const result = await evaluateWorkbook(book);
		expect(lengthOf(result.cells.map((cell) => cell.value))).toEqual([
			32767,
			...Array.from({ length: 15 }, () => VALUE),
		]);
	});

	it("is #VALUE! where & and REPLACE, UPPER and LOWER used to return the longer text", async () => {
		expect(
			await evaluateFormulas(
				[
					'REPT("a",32767)&REPT("a",32767)',
					'REPLACE(REPT("a",32767),1,0,REPT("b",32767))',
					'UPPER(REPT("ß",32767))',
					'LOWER(REPT("İ",32767))',
				],
				[],
			),
		).toEqual([VALUE, VALUE, VALUE, VALUE]);
	});

	it("is #VALUE! one character past the limit, from & and every built-in that builds text", async () => {
		expect(
			await evaluateFormulas(
				[
					'REPT("a",32767)&"b"',
					'CONCAT(REPT("a",32767),"b")',
					'CONCATENATE(REPT("a",32767),"b")',
					'TEXTJOIN("-",TRUE,REPT("a",32766),"b")',
					'SUBSTITUTE(REPT("a",32766)&"x","x","yy")',
					'SUBSTITUTE(REPT("a",32766)&"x","x","yy",1)',
					'REPLACE(REPT("a",32767),1,0,"b")',
					'UPPER(REPT("ß",16384))',
					'LOWER(REPT("İ",16384))',
					'REPT("a",32768)',
				],
				[],
			),
		).toEqual(Array.from({ length: 10 }, () => VALUE));
	});
});

describe("text of exactly 32,767 characters", () => {
	it("is still allowed from & and every built-in that builds text", async () => {
		const results = await evaluateFormulas(
			[
				'REPT("a",32766)&"b"',
				'CONCAT(REPT("a",32766),"b")',
				'CONCATENATE(REPT("a",32766),"b")',
				'TEXTJOIN("-",TRUE,REPT("a",32765),"b")',
				'SUBSTITUTE(REPT("a",32767),"a","b")',
				'SUBSTITUTE(REPT("a",32765)&"x","x","yy",1)',
				'REPLACE(REPT("a",32767),1,1,"b")',
				'UPPER(REPT("ß",16383)&"a")',
				'LOWER(REPT("İ",16383)&"a")',
				'REPT("a",32767)',
			],
			[],
		);
		expect(lengthOf(results)).toEqual(Array.from({ length: 10 }, () => 32767));
		expect(results[3]).toBe(`${"a".repeat(32765)}-b`);
		expect(results[7]).toBe(`${"SS".repeat(16383)}A`);
	});

	it("is still allowed from a range, the delimiters included", async () => {
		const rows = [["a".repeat(32760)], ["b".repeat(4)], ["c"]];
		expect(
			lengthOf(
				await evaluateFormulas(
					[
						"CONCAT(Data!A1:A3)",
						'TEXTJOIN(",",FALSE,Data!A1:A3)',
						'TEXTJOIN(",",TRUE,Data!A1:A3)',
					],
					rows,
				),
			),
		).toEqual([32765, 32767, 32767]);
	});
});

describe("REPT with an infinite count", () => {
	it("gives empty text for empty text, where it used to throw, and #VALUE! otherwise", async () => {
		// 1E+999 reads as infinity, as a literal and as a cell's value.
		expect(
			await evaluateFormulas(
				['REPT("",1E+999)', 'REPT("",Data!A1)', 'REPT("a",1E+999)', 'REPT("a",Data!A1)'],
				[[{ formula: "1E+999", value: 0 }]],
			),
		).toEqual(["", "", VALUE, VALUE]);
	});
});

describe("what decides between #VALUE! and another error", () => {
	it("is the order it was before: an error later in a range, or anywhere in TEXTJOIN, still wins", async () => {
		const rows = [["a".repeat(32767)], ["b"], [{ formula: "NA()", value: 0 }]];
		expect(
			await evaluateFormulas(
				[
					"CONCAT(Data!A1:A3)",
					'CONCAT(Data!A1,"b",NA())',
					'TEXTJOIN(",",TRUE,Data!A1,"b",NA())',
					'TEXTJOIN(",",TRUE,Data!A1:A3)',
					"Data!A1&NA()",
				],
				rows,
			),
		).toEqual([
			errorValue("#N/A"),
			VALUE,
			errorValue("#N/A"),
			errorValue("#N/A"),
			errorValue("#N/A"),
		]);
	});
});

describe("text longer than 32,767 characters that a reader stored", () => {
	// A1 holds 40,000 characters, which Excel cannot store but another producer can.
	const rows = [["a".repeat(40000)]];

	it("is returned as it is by a formula that passes it through, as before", async () => {
		expect(
			lengthOf(
				await evaluateFormulas(
					[
						"Data!A1",
						"IF(TRUE,Data!A1)",
						"PROPER(Data!A1)",
						"LEFT(Data!A1,40000)",
						"RIGHT(Data!A1,40000)",
						"MID(Data!A1,1,40000)",
						"TRIM(Data!A1)",
						"T(Data!A1)",
						"LEN(Data!A1)",
						'SUBSTITUTE(Data!A1,"zz","y")',
						'SUBSTITUTE(Data!A1,"","y")',
						'SUBSTITUTE(Data!A1,"zz","y",1)',
						'SUBSTITUTE(Data!A1,"","y",NA())',
					],
					rows,
				),
			),
		).toEqual(Array.from({ length: 13 }, () => 40000));
	});

	it("is #VALUE! once & or a built-in that builds text returns it", async () => {
		expect(
			await evaluateFormulas(
				[
					'Data!A1&""',
					"CONCAT(Data!A1)",
					'TEXTJOIN(",",TRUE,Data!A1)',
					'SUBSTITUTE(Data!A1,"a","a")',
					'SUBSTITUTE(Data!A1,"a","a",1)',
					'REPLACE(Data!A1,1,0,"")',
					"UPPER(Data!A1)",
					"LOWER(Data!A1)",
					"REPT(Data!A1,1)",
				],
				rows,
			),
		).toEqual(Array.from({ length: 9 }, () => VALUE));
	});

	it("can still be cut down to the limit", async () => {
		expect(
			lengthOf(
				await evaluateFormulas(
					[
						'SUBSTITUTE(Data!A1,"aaaaaaaaaa","")',
						'REPLACE(Data!A1,1,7233,"")',
						'LEFT(Data!A1,32767)&""',
					],
					rows,
				),
			),
		).toEqual([0, 32767, 32767]);
	});
});
