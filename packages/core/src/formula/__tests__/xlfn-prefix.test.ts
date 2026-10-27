import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import { evaluateWorkbook } from "../eval";
import { parseFormula } from "../parser";
import { errorValue, type ScalarValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.61, end to end: the `_xlfn.` names Excel stores for functions added after Excel 2007 evaluate
// like the bare names, and the AST keeps the stored text.

const data = [[10, 3]]; // Data!A1 = 10, Data!B1 = 3

const cases: readonly (readonly [formula: string, expected: ScalarValue])[] = [
	["IFNA(NA(),5)", 5],
	["IFS(FALSE,1,TRUE,2)", 2],
	['SWITCH(2,1,"a",2,"b")', "b"],
	["XOR(TRUE,FALSE,FALSE)", true],
	['CONCAT("a",Data!A1,"b")', "a10b"],
	['TEXTJOIN("-",TRUE,"a","","b")', "a-b"],
	["DAYS(Data!A1,Data!B1)", 7],
];

describe("_xlfn.-prefixed built-ins", () => {
	it.each(cases)("_xlfn.%s evaluates like the bare name", async (formula, expected) => {
		const results = await evaluateFormulas(
			[formula, `_xlfn.${formula}`, `_XLFN.${formula}`, `_xlfn._xlws.${formula}`],
			data,
		);
		expect(results).toEqual([expected, expected, expected, expected]);
	});

	it("keeps lazy arguments lazy", async () => {
		expect(
			await evaluateFormulas(["_xlfn.IFNA(1,NOW())", "_xlfn.IFS(TRUE,2,TRUE,NOW())"], data),
		).toEqual([1, 2]);
		await expect(evaluateFormulas(["_xlfn.IFNA(NA(),NOW())"], data)).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
	});

	it("checks the argument counts of the bare name", async () => {
		expect(
			await evaluateFormulas(
				["_xlfn.IFNA(1)", '_xlfn.TEXTJOIN("-",TRUE)', "_xlfn.DAYS(1,2,3)"],
				data,
			),
		).toEqual([errorValue("#VALUE!"), errorValue("#VALUE!"), errorValue("#VALUE!")]);
	});
});

describe("_xlfn. names that resolve to nothing", () => {
	it("stay #NAME?", async () => {
		const formulas = [
			"_xlfn.NOSUCH(1)",
			"_xlfn.XLOOKUP(1,Data!A1:A2,Data!B1:B2)",
			"_xlfn._xlfn.IFNA(NA(),5)",
			"_xlws.IFNA(NA(),5)",
		];
		expect(await evaluateFormulas(formulas, data)).toEqual(
			formulas.map(() => errorValue("#NAME?")),
		);
	});
});

describe("_xlfn. names and caller functions", () => {
	it("a caller function registered under the prefixed name wins; the bare name keeps the built-in", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							[
								{ formula: "_xlfn.IFNA(NA(),5)", value: 0 },
								{ formula: "IFNA(NA(),5)", value: 0 },
							],
						],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book, {
			functions: { "_xlfn.IFNA": { minArgs: 2, maxArgs: 2, evaluate: () => 99 } },
		});
		expect([result.get("S", "A1"), result.get("S", "B1")]).toEqual([99, 5]);
	});

	it("a prefixed name reaches a caller's override of the bare name", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[{ formula: "_xlfn.IFNA(NA(),5)", value: 0 }]] }],
			}),
		);
		const result = await evaluateWorkbook(book, {
			functions: { IFNA: { minArgs: 2, maxArgs: 2, evaluate: () => 7 } },
		});
		expect(result.get("S", "A1")).toBe(7);
	});

	it("gates a volatile function reached through the prefix", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[{ formula: "_xlfn.RANDARRAY()", value: 0 }]] }],
			}),
		);
		const functions = {
			RANDARRAY: { minArgs: 0, maxArgs: 5, volatile: true, evaluate: () => 0.5 },
		};
		await expect(evaluateWorkbook(book, { functions })).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
		const result = await evaluateWorkbook(book, { functions, random: () => 0.25 });
		expect(result.get("S", "A1")).toBe(0.5);
	});

	it("a prefix alone never reaches a function registered under the empty name", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							[
								{ formula: "_xlfn.()", value: 0 },
								{ formula: "_xlfn._xlws.()", value: 0 },
							],
						],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book, {
			functions: { "": { minArgs: 0, maxArgs: 0, evaluate: () => 3 } },
		});
		expect([result.get("S", "A1"), result.get("S", "B1")]).toEqual([
			errorValue("#NAME?"),
			errorValue("#NAME?"),
		]);
	});
});

describe("parseFormula of an _xlfn. call", () => {
	it("keeps the stored name", () => {
		expect(parseFormula("_xlfn.IFNA(NA(),5)")).toMatchObject({
			type: "call",
			name: "_xlfn.IFNA",
		});
		expect(parseFormula("_xlfn._xlws.IFNA(NA(),5)")).toMatchObject({
			type: "call",
			name: "_xlfn._xlws.IFNA",
		});
	});
});
