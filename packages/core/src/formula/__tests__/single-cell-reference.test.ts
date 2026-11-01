import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import type { DefinedName } from "../../ooxml/workbook";
import { openXlsx, type Workbook } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { eager, eagerByValue } from "../builtins";
import { evaluateCell, evaluateWorkbook, oneCellView, singleCellReference } from "../eval";
import type { EagerFunctionSpec } from "../functions";
import { parseFormula } from "../parser";
import { errorValue, isRangeView } from "../value";

// F11.86(a): a single-cell reference argument used to reach a function as the cell's value, so with
// D1 "n/a" and D2 2, SUM, AVERAGE and MAX of (D1,D2) gave #VALUE! where Excel gives 2. Every eager
// built-in except INDEX, VLOOKUP and HLOOKUP now receives a 1×1 view over the resolved cell, so it
// reads the reference exactly as the same cell written A1:A1 (D21). User functions, lazy built-ins and
// parameters that take one value are unchanged.

const f = (formula: string): CellInput => ({ formula, value: 0 });

// Sheet S: the cell under test in A1, criteria cells in B1:B3, a lookup table in M1:N3, and one
// formula per row in K. Every formula is evaluated, in order.
async function evaluateOnSheet(
	a1: CellInput,
	formulas: readonly string[],
	definedNames: readonly DefinedName[] = [],
) {
	const table: readonly (readonly CellInput[])[] = [
		[2, "two"],
		[3, "three"],
		["n/a", "text"],
	];
	const criteria: readonly CellInput[] = [2, "n/a", 3];
	const rows: CellInput[][] = [];
	for (let r = 0; r < Math.max(formulas.length, 3); r++) {
		const row: CellInput[] = Array(14).fill(null);
		if (r === 0) row[0] = a1;
		if (r < 3) {
			row[1] = criteria[r] ?? null;
			row[12] = table[r]?.[0] ?? null;
			row[13] = table[r]?.[1] ?? null;
		}
		const formula = formulas[r];
		if (formula !== undefined) row[10] = f(formula);
		rows.push(row);
	}
	const book = await openXlsx(
		await writeXlsx({
			sheets: [{ name: "S", rows }],
			...(definedNames.length > 0 ? { definedNames } : {}),
		}),
	);
	const result = await evaluateWorkbook(book);
	return formulas.map((_, i) => result.get("S", `K${i + 1}`));
}

// The workbook with `ref` left out of `sheetName`'s rows(), as a caller-built reader may list only
// the cells with a stored value; every other member reads through.
function withoutRowCell(book: Workbook, sheetName: string, ref: string): Workbook {
	return new Proxy(book, {
		get(target, key) {
			if (key !== "sheet") return Reflect.get(target, key, target);
			return (name: string) => {
				const sheet = target.sheet(name);
				if (name !== sheetName) return sheet;
				return new Proxy(sheet, {
					get(ws, member) {
						if (member === "rows") {
							return async function* () {
								for await (const row of ws.rows()) {
									yield {
										...row,
										cells: row.cells.filter((cell) => cell.ref !== ref),
									};
								}
							};
						}
						const value: unknown = Reflect.get(ws, member, ws);
						return typeof value === "function" ? value.bind(ws) : value;
					},
				});
			};
		},
	});
}

const CELL_NAME: readonly DefinedName[] = [{ name: "Cell", refersTo: "S!$A$1" }];

// Every kind of value a referenced cell can hold.
const KINDS: readonly (readonly [string, CellInput])[] = [
	["a number", 3],
	["text", "n/a"],
	["numeric text", "5"],
	["TRUE", true],
	["FALSE", false],
	["an empty string", ""],
	["a formula's empty string", f('IFERROR(1/0,"")')],
	["an error", f("NA()")],
	["nothing", null],
];

// The spellings of one cell reference; each must read as A1:A1 does.
const SPELLINGS = ["A1", "(A1)", "S!A1", "(S!A1)", "Cell", "(Cell)"];

// A template with its one `@` replaced by the argument's spelling.
function fill(template: string, spelling: string): string {
	const parts = template.split("@");
	if (parts.length !== 2) throw new Error(`template ${template} needs exactly one @`);
	return parts.join(spelling);
}

// Functions that read a reference argument as a range (@ is the argument).
const REFERENCE_FUNCTIONS = [
	"SUM(@,2)",
	"AVERAGE(@,2)",
	"COUNT(@,2)",
	"MIN(@,2)",
	"MAX(@,2)",
	"MEDIAN(@,2)",
	"LARGE(@,1)",
	"SMALL(@,1)",
	"AND(@,TRUE)",
	"OR(@,FALSE)",
	"XOR(@,TRUE)",
	"MATCH(2,@,0)",
	"COUNTA(@)",
	"COUNTBLANK(@)",
	'CONCAT(@,"|")',
];

// Parameters that take one value: unchanged, so each reads as the cell's plain value does. INDEX of a
// one-cell range returns that value itself, as HEAD passed a single-cell reference.
const VALUE_FUNCTIONS = [
	"ABS(@)",
	"LEN(@)",
	"N(@)",
	"T(@)",
	"VALUE(@)",
	"ISTEXT(@)",
	"ISNUMBER(@)",
	"ISBLANK(@)",
	"ISERROR(@)",
	"ISLOGICAL(@)",
	"NOT(@)",
	"ROUND(@,0)",
	"UPPER(@)",
	"LEFT(@,1)",
	'EXACT(@,"n/a")',
	"DATE(2020,@,1)",
	"COUNTIF(B1:B3,@)",
	"VLOOKUP(@,M1:N3,2,FALSE)",
	"MATCH(@,M1:M3,0)",
	"INDEX(M1:M3,@)",
	"IF(@,1,0)",
	"CHOOSE(1,@)",
	'IFERROR(@,"e")',
	'@&""',
];

describe("a single-cell reference reads as a reference in built-ins (F11.86(a))", () => {
	it("ignores text in a referenced cell, as Excel does", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							[null, null, null, "n/a"],
							[null, null, null, 2],
						],
					},
					{
						name: "R",
						rows: [[f("SUM(S!D1,S!D2)"), f("AVERAGE(S!D1,S!D2)"), f("MAX(S!D1,S!D2)")]],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect(["A1", "B1", "C1"].map((ref) => result.get("R", ref))).toEqual([2, 2, 2]);
	});

	it.each(KINDS)("reads every spelling of a cell holding %s as A1:A1", async (_, a1) => {
		const formulas = REFERENCE_FUNCTIONS.flatMap((template) => [
			fill(template, "A1:A1"),
			...SPELLINGS.map((spelling) => fill(template, spelling)),
		]);
		const values = await evaluateOnSheet(a1, formulas, CELL_NAME);
		const per = SPELLINGS.length + 1;
		for (let i = 0; i < REFERENCE_FUNCTIONS.length; i++) {
			const expected = values[i * per];
			for (let s = 1; s < per; s++) {
				expect(values[i * per + s], formulas[i * per + s]).toEqual(expected);
			}
		}
	});

	it.each(KINDS)(
		"keeps every one-value parameter as it was, for a cell holding %s",
		async (_, a1) => {
			const formulas = VALUE_FUNCTIONS.flatMap((template) => [
				fill(template, "INDEX(A1:A1,1,1)"),
				...SPELLINGS.map((spelling) => fill(template, spelling)),
			]);
			const values = await evaluateOnSheet(a1, formulas, CELL_NAME);
			const per = SPELLINGS.length + 1;
			for (let i = 0; i < VALUE_FUNCTIONS.length; i++) {
				const expected = values[i * per];
				for (let s = 1; s < per; s++) {
					expect(values[i * per + s], formulas[i * per + s]).toEqual(expected);
				}
			}
		},
	);

	it("keeps INDEX, VLOOKUP, HLOOKUP and SUMPRODUCT reading the cell's value, as 1.0.0 did", async () => {
		const values = await evaluateOnSheet(7, [
			"INDEX(A1,1)",
			"INDEX(A1,0)",
			"VLOOKUP(Z1,Z1,Z1,3)",
			"HLOOKUP(Z1,Z1,Z1,3)",
			"SUMPRODUCT(A1,{1,2,3})",
			"SUMPRODUCT((A1),{1,2,3})",
		]);
		expect(values).toEqual([7, 7, errorValue("#N/A"), errorValue("#N/A"), 7, 7]);
		// SUMPRODUCT keeps a referenced error ahead of its shape check (owner ruling, 2026-11-01).
		expect(await evaluateOnSheet(f("NA()"), ["SUMPRODUCT(A1,M1:M3)"])).toEqual([
			errorValue("#N/A"),
		]);
		expect(await evaluateOnSheet("n/a", ["SUMPRODUCT(A1,{1,2,3})"])).toEqual([0]);
		expect(await evaluateOnSheet(null, ["SUMPRODUCT(A1,{1,2,3})"])).toEqual([0]);
		const fixture = await openXlsx(await loadFixture("errors.xlsx"));
		expect(fixture.sheet("Feuil1").formula("A7")).toBe("VLOOKUP(B1,B1,B1,3)");
		expect(fixture.sheet("Feuil1").cell("A7")).toMatchObject({ type: "error", value: "#N/A" });
		expect(await evaluateCell(fixture, "Feuil1", "A7")).toEqual(errorValue("#N/A"));
	});

	it("gives D18's cases over an empty INDEX result Excel's answers", async () => {
		const values = await evaluateOnSheet(null, [
			'IFERROR(INDEX(A1:A2,2),"")',
			"SUM(K1)",
			"MAX(K1)",
			"AVERAGE(K1,2)",
		]);
		expect(values).toEqual(["", 0, 0, 2]);
	});

	it("passes a cell holding a range on as that range", async () => {
		const rows: CellInput[][] = [
			[f("G1:G3"), null, null, null, null, null, 1],
			[null, null, null, null, null, null, 2],
			[null, null, null, null, null, null, 3],
		];
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "S", rows },
					{ name: "R", rows: [[f("SUM(S!A1)"), f("COUNT(S!A1)")]] },
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect([result.get("R", "A1"), result.get("R", "B1")]).toEqual([6, 3]);
	});

	it("counts a formula cell its sheet's rows leave out as populated, through formula()", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "S", rows: [[f("2+3")]] },
					{
						name: "R",
						rows: [[f("COUNTA(S!A1)"), f("SUM(S!A1)"), f("COUNTBLANK(S!A1)")]],
					},
				],
			}),
		);
		// A reader whose rows() leaves S!A1 out, while formula("A1") still returns it.
		const result = await evaluateWorkbook(withoutRowCell(book, "S", "A1"));
		expect(["A1", "B1", "C1"].map((ref) => result.get("R", ref))).toEqual([1, 5, 0]);
	});

	it("agrees with SUMIFS and AVERAGEIFS on single-cell arguments (with F11.39)", async () => {
		const values = await evaluateOnSheet(null, [
			'SUMIF(G1,"",H1)',
			'SUMIFS(H1,G1,"")',
			'AVERAGEIF(G1,"",H1)',
			'AVERAGEIFS(H1,G1,"")',
		]);
		expect(values).toEqual([0, 0, errorValue("#DIV/0!"), errorValue("#DIV/0!")]);
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "S", rows: [[null, f("NA()")]] },
					{
						name: "R",
						rows: [
							[
								f('SUMIF(S!A1,"",S!B1)'),
								f('SUMIFS(S!B1,S!A1,"")'),
								f('AVERAGEIF(S!A1,"",S!B1)'),
								f('AVERAGEIFS(S!B1,S!A1,"")'),
							],
						],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect(["A1", "B1", "C1", "D1"].map((ref) => result.get("R", ref))).toEqual(
			Array(4).fill(errorValue("#N/A")),
		);
	});

	it("evaluates a reference that resolves to no cell as before", async () => {
		const names: readonly DefinedName[] = [
			{ name: "Rate", refersTo: "7" },
			{ name: "Label", refersTo: '"x"' },
			{ name: "Pair", refersTo: "S!$A$1:$A$2" },
		];
		const values = await evaluateOnSheet(
			4,
			[
				"SUM(Missing!A1)",
				"SUM(Nope)",
				"SUM(S:T!A1)",
				"SUM(Rate,1)",
				"SUM(Label)",
				"SUM(Pair)",
			],
			names,
		);
		expect(values).toEqual([
			errorValue("#REF!"),
			errorValue("#NAME?"),
			errorValue("#REF!"),
			8,
			errorValue("#VALUE!"),
			4,
		]);
	});
});

describe("a single-cell reference keeps its value for user functions (F8.2's contract)", () => {
	it("hands a user function, even one named SUM, the cell's value", async () => {
		const seen: unknown[] = [];
		const probe: EagerFunctionSpec = {
			minArgs: 1,
			maxArgs: 1,
			evaluate: ([arg]) => {
				seen.push(isRangeView(arg ?? null) ? "range" : arg);
				return 1;
			},
		};
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [["n/a", f("PROBE(A1)"), f("SUM(A1)")]] }],
			}),
		);
		await evaluateWorkbook(book, { functions: { PROBE: probe, SUM: probe } });
		expect(seen).toEqual(["n/a", "n/a"]);
	});
});

describe("references through the driver (F11.86(a))", () => {
	it("evaluates a 20,000-cell chain of single-cell references", async () => {
		const rows: CellInput[][] = [[1]];
		for (let r = 2; r <= 20_000; r++) rows.push([f(`SUM(A${r - 1},1)`)]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		expect(await evaluateCell(book, "S", "A20000")).toBe(20_000);
	}, 60_000);

	it("reports a cycle through single-cell references", async () => {
		const book = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[f("SUM(B1)"), f("SUM(A1,1)")]] }] }),
		);
		expect(await evaluateCell(book, "S", "A1")).toEqual(errorValue("#CYCLE!"));
	});

	it("spends one visit per formula cell, as before", async () => {
		const rows: CellInput[][] = [[1]];
		for (let r = 2; r <= 10; r++) rows.push([f(`SUM(A${r - 1},1)`)]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		expect(await evaluateCell(book, "S", "A10", { maxCellVisits: 9 })).toBe(10);
		await expect(evaluateCell(book, "S", "A10", { maxCellVisits: 8 })).rejects.toMatchObject({
			code: "budget-exceeded",
		});
	});
});

describe("singleCellReference", () => {
	it("returns a cell node, qualified or not", () => {
		expect(singleCellReference(parseFormula("A1"))).toEqual({ type: "cell", ref: "A1" });
		expect(singleCellReference(parseFormula("Sheet2!B3"))).toMatchObject({
			type: "cell",
			ref: "B3",
			sheet: { name: "Sheet2" },
		});
	});

	it("returns a name node", () => {
		expect(singleCellReference(parseFormula("Rate"))).toMatchObject({
			type: "name",
			name: "Rate",
		});
	});

	it("looks through parentheses", () => {
		expect(singleCellReference(parseFormula("((A1))"))).toEqual({ type: "cell", ref: "A1" });
		expect(singleCellReference(parseFormula("(Rate)"))).toMatchObject({ type: "name" });
	});

	it("is undefined for anything else", () => {
		for (const text of ["A1:A1", "A1+0", "1", '"A1"', "SUM(A1)", "{1}"]) {
			expect(singleCellReference(parseFormula(text)), text).toBeUndefined();
		}
	});
});

describe("oneCellView", () => {
	it("is a populated 1×1 view that returns the resolved value", () => {
		let checks = 0;
		const view = oneCellView({ sheet: "S", col: 4, row: 1 }, "n/a", () => {
			checks++;
			return true;
		});
		expect([view.sheet, view.startCol, view.startRow, view.endCol, view.endRow]).toEqual([
			"S",
			4,
			1,
			4,
			1,
		]);
		expect([...view.entries()]).toEqual([{ col: 4, row: 1, value: "n/a" }]);
		expect(view.cellAt(0, 0)).toBe("n/a");
		expect(view.populatedCount()).toBe(1);
		expect(checks).toBe(2);
	});

	it("lists nothing when the sheet holds nothing there", () => {
		const view = oneCellView({ sheet: "S", col: 1, row: 1 }, null, () => false);
		expect([...view.entries()]).toEqual([]);
		expect(view.cellAt(0, 0)).toBeNull();
		expect(view.cellCount - view.populatedCount()).toBe(1);
	});
});

describe("eager", () => {
	it("builds an eager built-in that gets single-cell references as 1×1 ranges", () => {
		const invoke = () => 1;
		expect(eager(1, 2, invoke)).toEqual({
			minArgs: 1,
			maxArgs: 2,
			volatile: false,
			lazyArgs: false,
			oneCellRanges: true,
			invoke,
		});
		expect(eager(0, 0, invoke, true).volatile).toBe(true);
	});
});

describe("eagerByValue", () => {
	it("builds an eager built-in that gets a single-cell reference's value", () => {
		const invoke = () => 1;
		expect(eagerByValue(2, 3, invoke)).toEqual({
			minArgs: 2,
			maxArgs: 3,
			volatile: false,
			lazyArgs: false,
			oneCellRanges: false,
			invoke,
		});
	});
});
