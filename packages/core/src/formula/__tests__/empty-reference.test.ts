import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.62: a formula whose result is an empty cell holds 0, as Excel shows it. Data!A1 is empty, and
// Data!B1:D1 reach it three ways: directly, through a lazy IF, and through CHOOSE.

const f = (formula: string): CellInput => ({ formula, value: 0 });
const indexData: readonly (readonly CellInput[])[] = [[null, null, "k", null, f("A1")], [7]];
const data: readonly (readonly CellInput[])[] = [
	[null, f("A1"), f("IF(TRUE,A1)"), f("CHOOSE(1,A1)")],
];

describe("a formula over an empty cell", () => {
	it("evaluates to 0, however it reaches the cell", async () => {
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "Data", rows: data }] }));
		const result = await evaluateWorkbook(book);
		expect(["B1", "C1", "D1"].map((ref) => result.get("Data", ref))).toEqual([0, 0, 0]);
		expect(await evaluateCell(book, "Data", "B1")).toBe(0);
	});

	it("is not blank to ISBLANK", async () => {
		expect(
			await evaluateFormulas(
				["ISBLANK(Data!B1)", "ISBLANK(Data!C1)", "ISBLANK(Data!D1)"],
				data,
			),
		).toEqual([false, false, false]);
	});

	it("counts as a number in COUNT, and not in COUNTBLANK", async () => {
		expect(
			await evaluateFormulas(
				["COUNT(Data!B1:D1)", "COUNTBLANK(Data!B1:D1)", "COUNTA(Data!B1:D1)"],
				data,
			),
		).toEqual([3, 0, 3]);
	});

	it("joins as 0 with &", async () => {
		expect(await evaluateFormulas(['Data!B1&"x"', 'Data!C1&"x"', 'Data!D1&"x"'], data)).toEqual(
			["0x", "0x", "0x"],
		);
	});

	it("compares as 0, not as an empty string", async () => {
		expect(
			await evaluateFormulas(
				['Data!B1=""', "Data!B1=0", "ISNUMBER(Data!B1)", "N(Data!B1)"],
				data,
			),
		).toEqual([false, true, true, 0]);
	});

	it("passes 0 on through a chain of formulas", async () => {
		expect(await evaluateFormulas(["Data!B1", 'Data!B1&"x"'], data)).toEqual([0, "0x"]);
	});

	it("passes 0 to a formula evaluated before the one it reads", async () => {
		// A1 and D1 come first in sheet order, so B1 and E1 are evaluated on their behalf: B1 on the
		// driver's frame stack, E1 through IF's lazy argument. C1 and F1 are empty.
		const rows = [[f('B1&"x"'), f("C1"), null, f('IF(TRUE,E1)&"x"'), f("F1")]];
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		const result = await evaluateWorkbook(book);
		expect(["A1", "B1", "D1", "E1"].map((ref) => result.get("S", ref))).toEqual([
			"0x",
			0,
			"0x",
			0,
		]);
	});

	it("holds the same 0 as an INDEX of the cell", async () => {
		expect(await evaluateFormulas(["Data!A1", "INDEX(Data!A1:A2,1)"], data)).toEqual([0, 0]);
	});

	it("evaluates to 0 through another sheet and a defined name", async () => {
		expect(
			await evaluateFormulas(["Data!Z9", "Empty"], data, [
				{ name: "Empty", refersTo: "Data!$A$1" },
			]),
		).toEqual([0, 0]);
	});

	it("evaluates to 0 when a caller function returns null", async () => {
		const book = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[f("MYNULL()")]] }] }),
		);
		const result = await evaluateWorkbook(book, {
			functions: { MYNULL: { minArgs: 0, maxArgs: 0, evaluate: () => null } },
		});
		expect(result.get("S", "A1")).toBe(0);
	});
});

describe("what F11.62 leaves as it was", () => {
	it("an empty cell that holds no formula stays empty", async () => {
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "Data", rows: data }] }));
		expect(await evaluateCell(book, "Data", "A1")).toBeNull();
	});

	it("an empty value inside a formula is still empty to its operators", async () => {
		expect(
			await evaluateFormulas(
				['Data!A1&"x"', 'Data!A1=""', "ISBLANK(Data!A1)", 'Empty&"x"'],
				data,
				[{ name: "Empty", refersTo: "Data!$A$1" }],
			),
		).toEqual(["x", true, true, "x"]);
	});

	it("an empty cell IF or CHOOSE returns is still empty inside the formula", async () => {
		// Only the cell's own final value becomes 0: IF and CHOOSE return the reference itself.
		expect(
			await evaluateFormulas(
				[
					'IF(TRUE,Data!A1)&"x"',
					'IF(TRUE,Data!A1)=""',
					"ISBLANK(IF(TRUE,Data!A1))",
					'CHOOSE(1,Data!A1)&"x"',
					"ISBLANK(CHOOSE(1,Data!A1))",
				],
				data,
			),
		).toEqual(["x", true, true, "x", true]);
	});

	it("a caller function's null is still empty inside the formula", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[f('MYNULL()&"x"'), f("ISBLANK(MYNULL())")]] }],
			}),
		);
		const result = await evaluateWorkbook(book, {
			functions: { MYNULL: { minArgs: 0, maxArgs: 0, evaluate: () => null } },
		});
		expect([result.get("S", "A1"), result.get("S", "B1")]).toEqual(["x", true]);
	});

	it("an empty string result stays an empty string", async () => {
		expect(await evaluateFormulas(['""', 'IF(TRUE,"")'], data)).toEqual(["", ""]);
	});

	it("a range result stays a range, so a formula over one empty cell still shows empty", async () => {
		// Recorded behavior: the scope keeps range results unchanged. Excel would show 0 here.
		expect(await evaluateFormulas(["Data!A1:A1", "SUM(Data!A1:A2)"], data)).toEqual([null, 0]);
	});
});

describe("IFERROR and IFNA over an empty cell", () => {
	// Microsoft: "If value or value_if_error is an empty cell, IFERROR treats it as an empty string
	// value ("")", and the same for IFNA. Data!A1 is empty, B1 and C1 hold IFERROR and IFNA over it,
	// D1 is 5, and E1 is a formula over A1, so it holds 0.
	const fallbackData: readonly (readonly CellInput[])[] = [
		[null, f("IFERROR(A1,1)"), f("IFNA(A1,1)"), 5, f("A1")],
	];

	it("evaluate to an empty string, for an empty value or an empty fallback", async () => {
		expect(
			await evaluateFormulas(
				[
					"IFERROR(Data!A1,1)",
					"IFNA(Data!A1,1)",
					"IFERROR(1/0,Data!A1)",
					"IFNA(NA(),Data!A1)",
					"IFERROR(Data!A1:A1,1)",
					"IFNA(Data!A1:A1,1)",
					"IFERROR(1/0,Data!A1:A1)",
					"IFNA(NA(),Data!A1:A1)",
				],
				fallbackData,
			),
		).toEqual(["", "", "", "", "", "", "", ""]);
	});

	it("still evaluate the fallback only on an error", async () => {
		// Without options.now, NOW() throws once evaluated, so these show which fallbacks ran.
		expect(
			await evaluateFormulas(
				["IFERROR(1,NOW())", "IFERROR(Data!A1,NOW())", "IFNA(Data!A1,NOW())"],
				fallbackData,
			),
		).toEqual([1, "", ""]);
		await expect(evaluateFormulas(["IFERROR(1/0,NOW())"], fallbackData)).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
	});

	it("are an empty string to the formulas that read them", async () => {
		expect(
			await evaluateFormulas(
				[
					"ISBLANK(Data!B1)",
					"ISBLANK(Data!C1)",
					"COUNTBLANK(Data!B1:C1)",
					"COUNTA(Data!B1:C1)",
					"COUNT(Data!B1:C1)",
					'Data!B1&"x"',
					'Data!B1=""',
					"Data!B1=0",
				],
				fallbackData,
			),
		).toEqual([false, false, 2, 2, 0, "x", true, false]);
	});

	it("are an empty string inside a formula too", async () => {
		expect(
			await evaluateFormulas(
				[
					"ISBLANK(IFERROR(Data!A1,1))",
					'IFERROR(Data!A1,1)&"x"',
					'IFNA(Data!A1,1)=""',
					"IFERROR(Data!A1,1)+1",
				],
				fallbackData,
			),
		).toEqual([false, "x", true, errorValue("#VALUE!")]);
	});

	it("keep every value that is not an empty cell", async () => {
		expect(
			await evaluateFormulas(
				[
					"IFERROR(Data!D1,1)",
					"IFERROR(Data!D1:D1,1)",
					"IFERROR(1/0,7)",
					"IFNA(1/0,1)",
					"IFERROR(0,1)",
					'IFERROR("",1)',
					"IFERROR(Data!E1,1)",
					"IFERROR(Data!E1:E1,1)",
				],
				fallbackData,
			),
		).toEqual([5, 5, 7, errorValue("#DIV/0!"), 0, "", 0, 0]);
	});
});

describe("INDEX of an empty cell", () => {
	// INDEX returns the cell itself, as Microsoft documents it returning a reference. indexData: A1 is
	// empty and A2 is 7; C1 is "k", D1 and C2 are empty; E1 is a formula over A1, so it holds 0.
	it("evaluates to 0 as a formula's final value, like the cell itself", async () => {
		expect(
			await evaluateFormulas(
				["INDEX(Data!A1:A2,1)", "INDEX(Data!A1:B2,1,1)", "INDEX(Data!A1,1)", "Data!A1"],
				indexData,
			),
		).toEqual([0, 0, 0, 0]);
	});

	it("is empty inside a formula, like the cell itself", async () => {
		expect(
			await evaluateFormulas(
				[
					'INDEX(Data!A1:A2,1)&"x"',
					'INDEX(Data!A1:B2,1,1)&"x"',
					'INDEX(Data!A1,1)&"x"',
					'INDEX(Data!A1:D1,4)&"x"',
					"ISBLANK(INDEX(Data!A1:A2,1))",
					'INDEX(Data!A1:A2,1)=""',
					"LEN(INDEX(Data!A1:A2,1))",
					"COUNT(INDEX(Data!A1:A2,1))",
					"COUNTA(INDEX(Data!A1:A2,1))",
					"IFERROR(INDEX(Data!A1:A2,1),1)",
					"IFNA(INDEX(Data!A1,1),1)",
				],
				indexData,
			),
		).toEqual(["x", "x", "x", "x", true, true, 0, 0, 0, "", ""]);
	});

	it("reads a formula target as the 0 it holds", async () => {
		expect(
			await evaluateFormulas(
				[
					'INDEX(Data!E1:E2,1)&"x"',
					"ISBLANK(INDEX(Data!E1:E2,1))",
					"COUNT(INDEX(Data!E1:E2,1))",
				],
				indexData,
			),
		).toEqual(["0x", false, 1]);
	});

	it("keeps a non-empty target as is", async () => {
		expect(
			await evaluateFormulas(['INDEX(Data!A1:A2,2)&"x"', "INDEX(Data!A1:A2,2)+1"], indexData),
		).toEqual(["7x", 8]);
	});
});

describe("recorded behavior, pending an owner ruling", () => {
	it("VLOOKUP and HLOOKUP read an empty target as 0 inside a formula", async () => {
		// Kept by the INDEX ruling: Microsoft documents no rule for an empty VLOOKUP or HLOOKUP target.
		// Excel idioms such as VLOOKUP(...)&"" read it as empty, yet ISBLANK(VLOOKUP(...)) is FALSE.
		expect(
			await evaluateFormulas(
				[
					'VLOOKUP("k",Data!C1:D1,2,FALSE)&"x"',
					'VLOOKUP("k",Data!C1:D1,2,TRUE)&"x"',
					'HLOOKUP("k",Data!C1:C2,2,FALSE)&"x"',
					'IFERROR(VLOOKUP("k",Data!C1:D1,2,FALSE),1)',
				],
				indexData,
			),
		).toEqual(["0x", "0x", "0x", 0]);
	});

	it("IFERROR and IFNA read an omitted argument as an empty cell, where IF and CHOOSE give 0", async () => {
		// The parser gives an omitted argument the empty value, so `IFERROR(1/0,)` meets the empty-cell
		// rule. Microsoft lists value_if_error and value_if_na as required and documents no result here.
		expect(
			await evaluateFormulas(
				["IFERROR(1/0,)", "IFNA(NA(),)", "IF(FALSE,1,)", "CHOOSE(2,1,)"],
				[[null]],
			),
		).toEqual(["", "", 0, 0]);
	});

	it("SUM over a single-cell reference to such an empty string gives #VALUE!, as over any text", async () => {
		// A single-cell reference argument acts like a literal (F11.86(a)); Excel ignores the text.
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [[null, f("IFERROR(A1,0)"), f("SUM(B1,1)"), f("SUM(B1:B1,1)")]],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect(["B1", "C1", "D1"].map((ref) => result.get("S", ref))).toEqual([
			"",
			errorValue("#VALUE!"),
			1,
		]);
	});
});
