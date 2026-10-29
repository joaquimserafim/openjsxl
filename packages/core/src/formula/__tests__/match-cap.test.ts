import { describe, expect, it } from "vitest";
import { threadCpuMs } from "../../ooxml/__tests__/complexity";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import type { CellInput } from "../../writer/types";
import {
	compileGlob,
	lookupEquals,
	lookupGlob,
	MAX_MATCH_TEXT,
	matchFn,
	parseCriteria,
	tableLookup,
} from "../builtins";
import { evaluateWorkbook } from "../eval";
import { ArrayView, errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.68 (D12): a criterion or a text lookup value longer than 255 characters is #VALUE! in every
// function that matches with it, and within the cap a wildcard COUNTIF over long cells is fast.

const VALUE = errorValue("#VALUE!");
const a = (n: number) => "a".repeat(n);

describe("MAX_MATCH_TEXT", () => {
	it("is 255, Excel's documented limit", () => {
		expect(MAX_MATCH_TEXT).toBe(255);
	});
});

describe("parseCriteria", () => {
	it("parses numbers, booleans and an empty cell without a cap", () => {
		expect(parseCriteria(errorValue("#N/A"))).toEqual(errorValue("#N/A"));
		expect(parseCriteria(5)).toEqual({ kind: "number", op: "=", operand: 5 });
		expect(parseCriteria(true)).toEqual({ kind: "boolean", op: "=", operand: true });
		expect(parseCriteria(null)).toEqual({ kind: "number", op: "=", operand: 0 });
	});

	it("parses text with its operator", () => {
		expect(parseCriteria(">5")).toEqual({ kind: "number", op: ">", operand: 5 });
		expect(parseCriteria("<>TRUE")).toEqual({ kind: "boolean", op: "<>", operand: true });
		expect(parseCriteria("false")).toEqual({ kind: "boolean", op: "=", operand: false });
		expect(parseCriteria(">b")).toEqual({ kind: "text", op: ">", operand: "b", glob: null });
		expect(parseCriteria("a*")).toEqual({
			kind: "text",
			op: "=",
			operand: "a*",
			glob: compileGlob("a*"),
		});
	});

	it("gives #VALUE! past 255 characters, the operator included", () => {
		expect(parseCriteria(a(255))).toMatchObject({ kind: "text", operand: a(255) });
		expect(parseCriteria(a(256))).toEqual(VALUE);
		expect(parseCriteria(`=${a(254)}`)).toMatchObject({ kind: "text", operand: a(254) });
		expect(parseCriteria(`=${a(255)}`)).toEqual(VALUE);
		expect(parseCriteria("1".repeat(256))).toEqual(VALUE);
	});
});

describe("lookupGlob and lookupEquals", () => {
	it("compiles a pattern only for text with a wildcard", () => {
		expect(lookupGlob("b*")).toEqual(compileGlob("b*"));
		expect(lookupGlob("b")).toBeNull();
		expect(lookupGlob(5)).toBeNull();
	});

	it("matches text by pattern, and anything else by value", () => {
		expect(lookupEquals(errorValue("#N/A"), "a", null)).toBe(false);
		expect(lookupEquals("Bat", "b*", lookupGlob("b*"))).toBe(true);
		expect(lookupEquals(5, "b*", lookupGlob("b*"))).toBe(false);
		expect(lookupEquals(true, "t*", lookupGlob("t*"))).toBe(false);
		expect(lookupEquals(null, "*", lookupGlob("*"))).toBe(false);
		expect(lookupEquals("a", "A", null)).toBe(true);
		expect(lookupEquals(1, true, null)).toBe(false);
	});
});

describe("tableLookup", () => {
	const table = new ArrayView([
		["a", 1],
		["b", 2],
	]);

	it("checks its arguments in order", () => {
		expect(tableLookup([errorValue("#N/A"), table, 2], "col")).toEqual(errorValue("#N/A"));
		expect(tableLookup(["a", "a", 2], "col")).toEqual(errorValue("#N/A"));
		expect(tableLookup(["a", table, "x"], "col")).toEqual(VALUE);
		expect(tableLookup(["a", table, 0], "col")).toEqual(VALUE);
		expect(tableLookup(["a", table, 3], "col")).toEqual(errorValue("#REF!"));
		expect(tableLookup(["a", table, 2, "x"], "col")).toEqual(VALUE);
	});

	it("matches approximately or exactly, with wildcards", () => {
		expect(tableLookup(["0", table, 2], "col")).toEqual(errorValue("#N/A"));
		expect(tableLookup(["bz", table, 2], "col")).toBe(2);
		expect(tableLookup(["B*", table, 2, false], "col")).toBe(2);
		expect(tableLookup(["c", table, 2, false], "col")).toEqual(errorValue("#N/A"));
	});

	it("gives #VALUE! for text past 255 characters, in either mode", () => {
		const long = new ArrayView([[a(256), 1]]);
		expect(tableLookup([a(256), long, 2, false], "col")).toEqual(VALUE);
		expect(tableLookup([a(256), long, 2], "col")).toEqual(VALUE);
		expect(tableLookup([a(255), new ArrayView([[a(255), 1]]), 2, false], "col")).toBe(1);
	});
});

describe("matchFn", () => {
	const list = new ArrayView([["a"], ["b"], ["c"]]);

	it("checks its arguments and its shape", () => {
		expect(matchFn([errorValue("#N/A"), list])).toEqual(errorValue("#N/A"));
		expect(matchFn(["a", list, "x"])).toEqual(VALUE);
		expect(matchFn(["a", "a"])).toBe(1);
		expect(matchFn(["a*", "ab"])).toBe(1);
		expect(matchFn(["a", "b"])).toEqual(errorValue("#N/A"));
		expect(
			matchFn([
				"a",
				new ArrayView([
					["a", "b"],
					["c", "d"],
				]),
			]),
		).toEqual(errorValue("#N/A"));
	});

	it("matches exactly with wildcards, or by order", () => {
		expect(matchFn(["b*", list, 0])).toBe(2);
		expect(matchFn(["z", list, 0])).toEqual(errorValue("#N/A"));
		expect(matchFn(["bb", list])).toBe(2);
		expect(matchFn(["0", list])).toEqual(errorValue("#N/A"));
		expect(matchFn(["bb", new ArrayView([["c"], ["b"], ["a"]]), -1])).toBe(1);
	});

	it("gives #VALUE! for text past 255 characters, in every mode", () => {
		for (const type of [0, 1, -1]) {
			expect(matchFn([a(256), new ArrayView([[a(256)]]), type])).toEqual(VALUE);
		}
		expect(matchFn([a(256), a(256)])).toEqual(VALUE);
		expect(matchFn([a(255), new ArrayView([[a(255)]]), 0])).toBe(1);
	});
});

describe("the nine functions that match", () => {
	// A1 holds 255 a's and A2 256; B holds 1 and 2. D1:E2 hold the same across, for HLOOKUP.
	const rows: CellInput[][] = [
		[a(255), 1, null, a(255), a(256)],
		[a(256), 2, null, 1, 2],
	];
	// The text is built with REPT, so no string literal in a formula needs to hold it.
	const functions = (text: string) => [
		`COUNTIF(Data!A1:A2,${text})`,
		`SUMIF(Data!A1:A2,${text},Data!B1:B2)`,
		`AVERAGEIF(Data!A1:A2,${text},Data!B1:B2)`,
		`COUNTIFS(Data!A1:A2,${text})`,
		`SUMIFS(Data!B1:B2,Data!A1:A2,${text})`,
		`AVERAGEIFS(Data!B1:B2,Data!A1:A2,${text})`,
		`VLOOKUP(${text},Data!A1:B2,2,FALSE)`,
		`HLOOKUP(${text},Data!D1:E2,2,FALSE)`,
		`MATCH(${text},Data!A1:A2,0)`,
	];

	it("still match with 255 characters", async () => {
		expect(await evaluateFormulas(functions('REPT("a",255)'), rows)).toEqual([
			1, 1, 1, 1, 1, 1, 1, 1, 1,
		]);
	});

	it("give #VALUE! with 256, even where a cell holds them", async () => {
		expect(await evaluateFormulas(functions('REPT("a",256)'), rows)).toEqual(
			Array.from({ length: 9 }, () => VALUE),
		);
	});

	it("count an operator in the criterion, and cap a lookup in every mode", async () => {
		expect(
			await evaluateFormulas(
				[
					'COUNTIF(Data!A1:A2,"*"&REPT("a",254))',
					'COUNTIF(Data!A1:A2,"*"&REPT("a",255))',
					'COUNTIF(Data!A1:A2,"<>"&REPT("a",253))',
					'COUNTIF(Data!A1:A2,"<>"&REPT("a",254))',
					'VLOOKUP(REPT("a",256),Data!A1:B2,2,TRUE)',
					'MATCH(REPT("a",256),Data!A1:A2,1)',
					'COUNTIF(Data!B1:B2,REPT("1",256))',
					"COUNTIF(Data!B1:B2,1)",
				],
				rows,
			),
		).toEqual([2, VALUE, 2, VALUE, VALUE, VALUE, VALUE, 1]);
	});

	it("read a capped value first, as they read an error in it", async () => {
		expect(
			await evaluateFormulas(
				[
					'VLOOKUP(REPT("a",256),Data!A1:B2,3,FALSE)',
					"VLOOKUP(NA(),Data!A1:B2,3,FALSE)",
					'MATCH(REPT("a",256),Data!A1:A2,NA())',
					'COUNTIFS(Data!A1:A2,REPT("a",256),Data!A1:A2,NA())',
					"COUNTIFS(Data!A1:A2,NA(),Data!A1:A2,1/0)",
				],
				rows,
			),
		).toEqual([VALUE, errorValue("#N/A"), VALUE, VALUE, errorValue("#N/A")]);
	});

	it("keep the ?, * and ~ semantics", async () => {
		const words: CellInput[][] = [["abc"], ["a*c"], ["a?c"], ["axyc"], ["A~C"], ["ß"]];
		expect(
			await evaluateFormulas(
				[
					'COUNTIF(Data!A1:A6,"a?c")',
					'COUNTIF(Data!A1:A6,"a*c")',
					'COUNTIF(Data!A1:A6,"a~*c")',
					'COUNTIF(Data!A1:A6,"a~?c")',
					'COUNTIF(Data!A1:A6,"a~c*")',
					'COUNTIF(Data!A1:A6,"<>a*")',
					'COUNTIF(Data!A1:A6,"*ß*")',
					'COUNTIF(Data!A1:A6,"s?")',
					'MATCH("A?C",Data!A1:A6,0)',
					'VLOOKUP("*y*",Data!A1:A6,1,FALSE)',
				],
				words,
			),
		).toEqual([4, 5, 1, 1, 1, 1, 0, 1, 1, "axyc"]);
	});
});

describe("a wildcard COUNTIF over long cells", () => {
	it("takes within 50 times a plain-text COUNTIF over the same 1,024 cells of 32,767 characters", async () => {
		const cells = 1024;
		const rows = Array.from({ length: cells }, (): CellInput[] => [
			{ formula: 'REPT("a",32767)', value: 0 },
		]);
		// The CPU time to evaluate one COUNTIF over the cells with this criterion.
		const time = async (criterion: string) => {
			const book = await openXlsx(
				await writeXlsx({
					sheets: [
						{ name: "Data", rows },
						{
							name: "R",
							rows: [
								[
									{
										formula: `COUNTIF(Data!A1:A${cells},"${criterion}")`,
										value: 0,
									},
								],
							],
						},
					],
				}),
			);
			await evaluateWorkbook(book);
			const start = threadCpuMs();
			const result = await evaluateWorkbook(book);
			expect(result.get("R", "A1")).toBe(0);
			return threadCpuMs() - start;
		};
		const plain = await time("zzz");
		// The last pattern's run, with a mismatch near its start, defeats indexOf's skip tables.
		for (const pattern of [
			`*${"a?".repeat(126)}b*`,
			`*${"?".repeat(252)}b*`,
			`*${a(252)}b*`,
			`*ab${a(251)}*`,
		]) {
			expect(pattern).toHaveLength(255);
			expect(await time(pattern)).toBeLessThan(plain * 50);
		}
	}, 300_000);
});
