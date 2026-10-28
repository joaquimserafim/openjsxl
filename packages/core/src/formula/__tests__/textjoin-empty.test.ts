import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import type { CellInput } from "../../writer/types";
import { joinKeepingEmpty, textjoin } from "../builtins";
import { evaluateCell } from "../eval";
import { ArrayView, type EvalValue, errorValue, RangeView } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.97: with ignore_empty FALSE, TEXTJOIN gives one empty item for every empty cell in a range, as
// Excel does, counted from the range's bounds so the cost follows the populated cells.

const VALUE = errorValue("#VALUE!");
const bold: CellInput = { value: null, style: { font: { bold: true } } };

// A column A1:A{height} whose populated rows hold `value(row)`, counting how many cells are read.
function column(height: number, rows: readonly number[], value: (row: number) => EvalValue) {
	const reads: number[] = [];
	const view = new RangeView(
		"S",
		1,
		1,
		1,
		height,
		() => rows.map((row): readonly [number, number] => [1, row]),
		(_, row) => {
			reads.push(row);
			return value(row);
		},
	);
	return { view, reads };
}

describe("TEXTJOIN with ignore_empty FALSE", () => {
	it("gives an empty item for an empty cell inside a range, styled or absent", async () => {
		const formula = ['TEXTJOIN(",",FALSE,Data!A1:A3)'];
		expect(await evaluateFormulas(formula, [["a"], [], ["c"]])).toEqual(["a,,c"]);
		expect(await evaluateFormulas(formula, [["a"], [bold], ["c"]])).toEqual(["a,,c"]);
	});

	it("counts the empty cells before the first populated cell and after the last", async () => {
		expect(
			await evaluateFormulas(
				['TEXTJOIN(",",FALSE,Data!B1:B3)', 'TEXTJOIN(",",FALSE,Data!A1:A5)'],
				[["a"], [null, "b"], ["c"]],
			),
		).toEqual([",b,", "a,,c,,"]);
	});

	it("walks a 2-D range in row-major order", async () => {
		expect(
			await evaluateFormulas(
				['TEXTJOIN(",",FALSE,Data!B1:C2)'],
				[
					[null, "x"],
					[null, null, "y"],
				],
			),
		).toEqual(["x,,,y"]);
	});

	it("allows 32,767 characters of delimiters over empty cells and refuses one more", async () => {
		const [full, over] = await evaluateFormulas(
			['TEXTJOIN(",",FALSE,Data!C1:C32768)', 'TEXTJOIN(",",FALSE,Data!C1:C32769)'],
			[["a"]],
		);
		expect(full).toBe(",".repeat(32767));
		expect(over).toEqual(VALUE);
	});

	it("is #VALUE! over a whole column, whose empty cells pass the limit", async () => {
		expect(
			await evaluateFormulas(['TEXTJOIN(",",FALSE,Data!A:A)'], [["a"], [], ["c"]]),
		).toEqual([VALUE]);
	});

	it("joins the populated cells of a whole sheet with an empty delimiter, at their cost", async () => {
		// A1, C5 and Z100.
		const rows = Array.from({ length: 100 }, (_, i): CellInput[] =>
			i === 0
				? ["a"]
				: i === 4
					? [null, null, "b"]
					: i === 99
						? [...Array(25).fill(null), "c"]
						: [],
		);
		expect(await evaluateFormulas(['TEXTJOIN("",FALSE,Data!A1:XFD1048576)'], rows)).toEqual([
			"abc",
		]);
	});

	it("keeps an empty item for every empty cell across several ranges and scalars", async () => {
		expect(
			await evaluateFormulas(
				[
					'TEXTJOIN(",",FALSE,Data!A1:A2,"m",Data!C1:C2)',
					'TEXTJOIN("-",FALSE,"a","","b",Data!B1)',
				],
				[
					["a", null],
					[null, null, "d"],
				],
			),
		).toEqual(["a,,m,,d", "a--b-"]);
	});

	it("gives #VALUE! once the limit passes, before an error in a range cell it no longer reads", async () => {
		// A1 holds "a" and A40000 #N/A, with 39,998 empty cells between them; B1:B3 hold 32,767
		// characters, "b" and #N/A, with no gap.
		const rows = Array.from({ length: 40000 }, (_, i): CellInput[] =>
			i === 0
				? ["a", { formula: 'REPT("a",32767)', value: 0 }]
				: i === 1
					? [null, "b"]
					: i === 2
						? [null, { formula: "NA()", value: 0 }]
						: i === 39999
							? [{ formula: "NA()", value: 0 }]
							: [],
		);
		expect(
			await evaluateFormulas(
				[
					'TEXTJOIN(",",FALSE,Data!A1:A40000)',
					'TEXTJOIN(",",TRUE,Data!A1:A40000)',
					'TEXTJOIN(",",FALSE,Data!B1:B3)',
					'TEXTJOIN(",",TRUE,Data!B1:B3)',
				],
				rows,
			),
		).toEqual([VALUE, errorValue("#N/A"), VALUE, errorValue("#N/A")]);
	});

	it("still lets a scalar's error beat the limit, as with TRUE", async () => {
		expect(
			await evaluateFormulas(
				[
					'TEXTJOIN(",",FALSE,Data!A1:A2,NA())',
					'TEXTJOIN(",",FALSE,REPT("a",32767),"b",NA())',
					'TEXTJOIN(",",TRUE,REPT("a",32767),"b",NA())',
					'TEXTJOIN(",",FALSE,Data!A:A,NA())',
					'TEXTJOIN(",",FALSE,Data!A:A,"x")',
				],
				[["a"], [], ["c"]],
			),
		).toEqual([
			errorValue("#N/A"),
			errorValue("#N/A"),
			errorValue("#N/A"),
			errorValue("#N/A"),
			VALUE,
		]);
	});

	it("gives #VALUE! for a cycle through a cell past the limit, which it never reads", async () => {
		// A1 joins B1:B40000, and B40000 refers back to A1.
		const rows = Array.from({ length: 40000 }, (_, i): CellInput[] =>
			i === 0
				? [
						{ formula: 'TEXTJOIN(",",FALSE,B1:B40000)', value: 0 },
						"a",
						{ formula: 'TEXTJOIN(",",TRUE,D1:D40000)', value: 0 },
						"a",
					]
				: i === 39999
					? [null, { formula: "A1", value: 0 }, null, { formula: "C1", value: 0 }]
					: [],
		);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		expect(await evaluateCell(book, "S", "A1")).toEqual(VALUE);
		expect(await evaluateCell(book, "S", "C1")).toEqual(errorValue("#CYCLE!"));
	});
});

describe("TEXTJOIN with ignore_empty TRUE", () => {
	it("gives the results it gave before", async () => {
		const rows = [["a"], [], ["c"], [bold], [""]];
		expect(
			await evaluateFormulas(
				[
					'TEXTJOIN(",",TRUE,Data!A1:A5)',
					'TEXTJOIN(",",TRUE,Data!A:A)',
					'TEXTJOIN(",",TRUE,Data!A1:XFD1048576)',
					'TEXTJOIN("-",TRUE,"a","","b")',
				],
				rows,
			),
		).toEqual(["a,c", "a,c", "a,c", "a-b"]);
	});
});

describe("textjoin", () => {
	it("hands ignore_empty FALSE to joinKeepingEmpty, after the delimiter and flag are read", () => {
		const { view } = column(3, [2], () => "b");
		expect(textjoin([",", false, view])).toBe(",b,");
		expect(textjoin([",", true, view])).toBe("b");
		expect(textjoin([errorValue("#N/A"), false, view])).toEqual(errorValue("#N/A"));
		expect(textjoin([",", "yes", view])).toEqual(VALUE);
	});
});

describe("joinKeepingEmpty", () => {
	it("joins scalars as items, an empty one included", () => {
		expect(joinKeepingEmpty(",", ["a", null, "b", 1, true])).toBe("a,,b,1,TRUE");
		expect(joinKeepingEmpty(",", [])).toBe("");
	});

	it("gives a range one item per cell, empty cells counted from its bounds", () => {
		expect(joinKeepingEmpty(",", [column(3, [2], () => "b").view])).toBe(",b,");
		expect(joinKeepingEmpty(",", [column(5, [3, 1], (row) => `r${row}`).view])).toBe(
			"r1,,r3,,",
		);
		expect(joinKeepingEmpty(",", [column(2, [], () => "x").view])).toBe(",");
		expect(joinKeepingEmpty(",", [column(1, [], () => "x").view])).toBe("");
		const grid = new RangeView(
			"S",
			2,
			1,
			3,
			2,
			() => [
				[3, 2],
				[2, 1],
			],
			(col, row) => `${col}${row}`,
		);
		expect(joinKeepingEmpty(",", [grid])).toBe("21,,,32");
		expect(
			joinKeepingEmpty(",", [
				new ArrayView([
					["a", "b"],
					["c", "d"],
				]),
			]),
		).toBe("a,b,c,d");
	});

	it("propagates an error from a scalar or from a cell", () => {
		expect(joinKeepingEmpty(",", ["a", errorValue("#N/A")])).toEqual(errorValue("#N/A"));
		expect(joinKeepingEmpty(",", [column(3, [2], () => errorValue("#DIV/0!")).view])).toEqual(
			errorValue("#DIV/0!"),
		);
	});

	it("counts every delimiter: exactly 32,767 characters is allowed, one more is not", () => {
		expect(joinKeepingEmpty(",", [column(32768, [], () => "x").view])).toBe(",".repeat(32767));
		expect(joinKeepingEmpty(",", [column(32769, [], () => "x").view])).toEqual(VALUE);
		expect(joinKeepingEmpty("--", [column(16384, [], () => "x").view])).toBe("-".repeat(32766));
		expect(joinKeepingEmpty("--", ["a", column(16384, [], () => "x").view])).toEqual(VALUE);
		expect(joinKeepingEmpty(",", [column(32767, [32767], () => "x").view])).toBe(
			`${",".repeat(32766)}x`,
		);
		expect(joinKeepingEmpty(",", [column(32767, [32767], () => "xy").view])).toEqual(VALUE);
		expect(joinKeepingEmpty("-", ["a".repeat(32767), "b"])).toEqual(VALUE);
		expect(joinKeepingEmpty("-", ["a".repeat(32766), ""])).toBe(`${"a".repeat(32766)}-`);
	});

	it("reads no cell once the limit passes", () => {
		const gap = column(40000, [1, 40000], () => "x");
		expect(joinKeepingEmpty(",", [gap.view])).toEqual(VALUE);
		expect(gap.reads).toEqual([1]);
		const long = column(3, [1, 2, 3], () => "a".repeat(20000));
		expect(joinKeepingEmpty("", [long.view])).toEqual(VALUE);
		expect(long.reads).toEqual([1, 2]);
		const after = column(3, [1], () => errorValue("#N/A"));
		expect(joinKeepingEmpty("", ["a".repeat(32767), "b", after.view])).toEqual(VALUE);
		expect(after.reads).toEqual([]);
	});

	it("still reads every scalar past the limit, so a scalar's error wins", () => {
		const after = column(3, [1], () => "x");
		expect(
			joinKeepingEmpty("", ["a".repeat(32767), "b", after.view, errorValue("#N/A")]),
		).toEqual(errorValue("#N/A"));
		expect(after.reads).toEqual([]);
		expect(
			joinKeepingEmpty(",", [column(40000, [], () => "x").view, errorValue("#DIV/0!")]),
		).toEqual(errorValue("#DIV/0!"));
		expect(joinKeepingEmpty("-", ["a".repeat(32767), "b", "c"])).toEqual(VALUE);
	});

	it("does not count empty cells with an empty delimiter, where they add nothing", () => {
		const sheet = new RangeView(
			"S",
			1,
			1,
			16384,
			1048576,
			() => [[3, 7]],
			() => "x",
		);
		expect(joinKeepingEmpty("", [sheet])).toBe("x");
		expect(joinKeepingEmpty("", [column(40000, [1, 40000], (row) => `${row}`).view])).toBe(
			"140000",
		);
		expect(joinKeepingEmpty(",", [sheet])).toEqual(VALUE);
	});

	it("costs time in proportion to the populated cells of a whole-sheet range", () => {
		const ratio = bestDoublingRatio((n) => {
			const rows = Array.from({ length: n }, (_, i) => n * 2 - i);
			const view = new RangeView(
				"S",
				1,
				1,
				16384,
				1048576,
				() => rows.map((row): readonly [number, number] => [1 + (row % 16384), row]),
				() => "",
			);
			return () => textjoin(["", false, view]);
		}, 20_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);
});
