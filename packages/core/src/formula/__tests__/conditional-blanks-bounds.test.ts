import { describe, expect, it, vi } from "vitest";
import { MAX_COL, MAX_ROW } from "../../ooxml/a1";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { BUILTIN_ENTRIES } from "../builtins";
import { evaluateCell } from "../eval";
import type { EvalContext } from "../functions";
import { type EvalValue, errorValue, RangeView } from "../value";

const context: EvalContext = {
	date1904: false,
	now: () => 0,
	random: () => 0,
};

type PopulatedCell = readonly [col: number, row: number, value: EvalValue];

function invoke(name: string, args: readonly EvalValue[]): EvalValue {
	const spec = BUILTIN_ENTRIES.find(([key]) => key === name)?.[1];
	if (spec === undefined || spec.lazyArgs) throw new Error(`missing eager builtin ${name}`);
	return spec.invoke(args, context);
}

// A genuine full-size RangeView with just a few populated coordinates. The hard read budget
// aborts a dense scan immediately; assertions below also cap sparse iteration and resolution work.
function sparseRange(width: number, cells: readonly PopulatedCell[] = []) {
	const values = new Map(cells.map(([col, row, value]) => [`${col},${row}`, value]));
	const coordinates: [number, number][] = cells.map(([col, row]) => [col, row]);
	const populated = vi.fn(() => coordinates);
	let reads = 0;
	const resolve = vi.fn((col: number, row: number): EvalValue => {
		if (++reads > 128) throw new Error("conditional aggregate exceeded its sparse read budget");
		return values.get(`${col},${row}`) ?? null;
	});
	const range = new RangeView("S", 1, 1, width, MAX_ROW, populated, resolve);
	return { range, populated, resolve };
}

const modes: readonly [string, number][] = [
	["SUMIFS", 7],
	["AVERAGEIFS", 3.5],
];

describe("multi-criterion blank matching stays sparse (F11.12)", () => {
	it.each([
		1,
		MAX_COL,
	])("counts an empty %i-column rectangle arithmetically without resolving blank cells", (width) => {
		const first = sparseRange(width);
		const second = sparseRange(width);
		expect(invoke("COUNTIFS", [first.range, "", second.range, ""])).toBe(width * MAX_ROW);
		expect(first.resolve).not.toHaveBeenCalled();
		expect(second.resolve).not.toHaveBeenCalled();
		expect(
			first.populated.mock.calls.length + second.populated.mock.calls.length,
		).toBeLessThanOrEqual(4);
	});

	it("deduplicates overlapping populated offsets while preserving exact full-grid counts", () => {
		const first = sparseRange(MAX_COL, [
			[1, 1, "x"],
			[MAX_COL, MAX_ROW, ""],
		]);
		const second = sparseRange(MAX_COL, [
			[1, 1, "x"],
			[6, 6, "y"],
			[MAX_COL, MAX_ROW, ""],
		]);
		// Three distinct populated offsets: two fail, the explicit empty-string position matches.
		// Both ranges overlap at the first and last cells, including offset 2^34 - 1.
		expect(invoke("COUNTIFS", [first.range, "", second.range, ""])).toBe(MAX_COL * MAX_ROW - 2);
		expect(
			first.populated.mock.calls.length + second.populated.mock.calls.length,
		).toBeLessThanOrEqual(8);
		expect(
			first.resolve.mock.calls.length + second.resolve.mock.calls.length,
		).toBeLessThanOrEqual(32);
	});

	it("retains a blank-excluding sparse driver without enumerating the other criterion", () => {
		const blanks = sparseRange(MAX_COL);
		blanks.populated.mockImplementation(() => {
			throw new Error("a companion criterion must not be enumerated");
		});
		const driver = sparseRange(MAX_COL, [[MAX_COL, MAX_ROW, 1]]);
		expect(invoke("COUNTIFS", [blanks.range, "<>x", driver.range, ">0"])).toBe(1);
		expect(blanks.populated).not.toHaveBeenCalled();
		expect(
			blanks.resolve.mock.calls.length + driver.resolve.mock.calls.length,
		).toBeLessThanOrEqual(8);
	});

	it.each(
		modes,
	)("%s visits populated values even when both criteria rectangles are entirely blank", (name, expected) => {
		const values = sparseRange(MAX_COL, [
			[1, 1, 3],
			[MAX_COL, MAX_ROW, 4],
		]);
		const first = sparseRange(MAX_COL);
		const second = sparseRange(MAX_COL);
		for (const criterion of [first, second]) {
			criterion.populated.mockImplementation(() => {
				throw new Error("value-driven aggregation must not enumerate criteria ranges");
			});
		}
		expect(invoke(name, [values.range, first.range, "", second.range, ""])).toBe(expected);
		expect(first.populated).not.toHaveBeenCalled();
		expect(second.populated).not.toHaveBeenCalled();
		expect(values.populated.mock.calls.length).toBeLessThanOrEqual(4);
		expect(
			values.resolve.mock.calls.length +
				first.resolve.mock.calls.length +
				second.resolve.mock.calls.length,
		).toBeLessThanOrEqual(16);
	});

	it.each([
		"SUMIFS",
		"AVERAGEIFS",
	])("%s never resolves a populated value excluded by its criteria", (name) => {
		const values = sparseRange(MAX_COL, [
			[1, 1, 3],
			[MAX_COL, MAX_ROW, 4],
		]);
		values.resolve.mockImplementation((col, row) => {
			if (col === MAX_COL && row === MAX_ROW) {
				throw new Error("an excluded value must remain unevaluated");
			}
			return 3;
		});
		const first = sparseRange(MAX_COL, [[MAX_COL, MAX_ROW, "excluded"]]);
		const second = sparseRange(MAX_COL);
		expect(invoke(name, [values.range, first.range, "", second.range, ""])).toBe(3);
		expect(values.resolve.mock.calls).not.toContainEqual([MAX_COL, MAX_ROW]);
	});

	it("collects COUNTIFS union positions without evaluating an already-excluded companion", () => {
		const first = sparseRange(MAX_COL, [[1, 1, "excluded"]]);
		const second = sparseRange(MAX_COL, [[1, 1, "unevaluated"]]);
		second.resolve.mockImplementation(() => {
			throw new Error("an earlier criterion excluded this populated position");
		});
		expect(invoke("COUNTIFS", [first.range, "", second.range, ""])).toBe(MAX_COL * MAX_ROW - 1);
		expect(second.resolve).not.toHaveBeenCalled();
	});

	it.each([
		"SUMIFS",
		"AVERAGEIFS",
	])("%s checks a later blank-excluding criterion before an earlier unreadable criterion", (name) => {
		const values = sparseRange(MAX_COL, [[1, 1, 3]]);
		const earlier = sparseRange(MAX_COL);
		earlier.resolve.mockImplementation(() => {
			throw new Error("the blank-excluding criterion must reject this position first");
		});
		const driver = sparseRange(MAX_COL);
		expect(invoke(name, [values.range, earlier.range, "", driver.range, ">0"])).toEqual(
			name === "SUMIFS" ? 0 : errorValue("#DIV/0!"),
		);
		expect(earlier.resolve).not.toHaveBeenCalled();
		expect(values.resolve).not.toHaveBeenCalled();
	});

	it.each([
		"COUNTIFS",
		"SUMIFS",
		"AVERAGEIFS",
	])("%s preserves the RangeView receiver of a populated-coordinate callback", (name) => {
		const range = new RangeView(
			"S",
			2,
			3,
			2,
			MAX_ROW,
			function (this: RangeView): Iterable<readonly [number, number]> {
				return [[this.startCol, this.startRow]];
			},
			() => 7,
		);
		const args =
			name === "COUNTIFS" ? [range, ">0", range, ">0"] : [range, range, ">0", range, ">0"];
		expect(invoke(name, args)).toBe(name === "COUNTIFS" ? 1 : 7);
	});

	it("keeps public entries evaluation attached to pulling an entry", () => {
		const values = sparseRange(MAX_COL, [[1, 1, 3]]);
		const entries = values.range.entries();
		expect(values.resolve).not.toHaveBeenCalled();
		const first = entries.next();
		expect(values.resolve).toHaveBeenCalledOnce();
		expect(first).toEqual({ done: false, value: { col: 1, row: 1, value: 3 } });
	});

	it.each([
		"COUNTIFS",
		"SUMIFS",
		"AVERAGEIFS",
	])("%s accepts a transparent RangeView proxy returned by a caller function", async (name) => {
		const raw = new RangeView(
			"Data",
			1,
			1,
			1,
			2,
			() => [[1, 1]],
			() => 1,
		);
		const proxy = new Proxy(raw, {});
		const formula =
			name === "COUNTIFS"
				? 'COUNTIFS(PROXY(),">0",PROXY(),">0")'
				: `${name}(PROXY(),PROXY(),">0",PROXY(),">0")`;
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "Results", rows: [[{ formula, value: 1 }]] }],
			}),
		);
		await expect(
			evaluateCell(book, "Results", "A1", {
				functions: { PROXY: { minArgs: 0, maxArgs: 0, evaluate: () => proxy } },
			}),
		).resolves.toBe(1);
	});

	it.each([
		"SUMIFS",
		"AVERAGEIFS",
	])("%s handles an empty value rectangle without resolving all-blank criteria", (name) => {
		const values = sparseRange(MAX_COL);
		const first = sparseRange(MAX_COL);
		const second = sparseRange(MAX_COL);
		expect(invoke(name, [values.range, first.range, "", second.range, ""])).toEqual(
			name === "SUMIFS" ? 0 : errorValue("#DIV/0!"),
		);
		expect(values.resolve).not.toHaveBeenCalled();
		expect(first.resolve).not.toHaveBeenCalled();
		expect(second.resolve).not.toHaveBeenCalled();
	});

	it.each([
		"COUNTIFS",
		"SUMIFS",
		"AVERAGEIFS",
	])("%s still enforces evaluator fuel when its sparse path visits formula cells", async (name) => {
		const rows: CellInput[][] = Array.from({ length: 5 }, (_, index) =>
			name === "COUNTIFS"
				? [{ formula: '""', value: "" }]
				: [null, null, { formula: `${index}+1`, value: index + 1 }],
		);
		const formula =
			name === "COUNTIFS"
				? 'COUNTIFS(Data!A:A,"",Data!B:B,"")'
				: `${name}(Data!C:C,Data!A:A,"",Data!B:B,"")`;
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Data", rows },
					{ name: "Results", rows: [[{ formula, value: 0 }]] },
				],
			}),
		);
		await expect(
			evaluateCell(book, "Results", "A1", { maxCellVisits: 2 }),
		).rejects.toMatchObject({ name: "FormulaError", code: "budget-exceeded" });
		await expect(evaluateCell(book, "Results", "A1", { maxCellVisits: 64 })).resolves.toBe(
			name === "COUNTIFS" ? MAX_ROW : name === "SUMIFS" ? 15 : 3,
		);
	});

	it.each([
		"SUMIFS",
		"AVERAGEIFS",
	])("%s preserves lazy exclusion of formula values under a one-visit fuel budget", async (name) => {
		for (const excludedFormula of ["NOW()", "2+2"]) {
			const book = await openXlsx(
				await writeXlsx({
					sheets: [
						{
							name: "Data",
							rows: [
								[1, 3],
								[0, { formula: excludedFormula, value: 0 }],
							],
						},
						{
							name: "Results",
							rows: [
								[
									{
										formula: `${name}(Data!B:B,Data!A:A,">0",Data!A:A,">0")`,
										value: 3,
									},
								],
							],
						},
					],
				}),
			);
			// NOW is deliberately unconfigured. Neither it nor the excluded arithmetic formula
			// may run: only the Results cell itself consumes the available formula-visit fuel.
			await expect(evaluateCell(book, "Results", "A1", { maxCellVisits: 1 })).resolves.toBe(
				3,
			);
		}
	});
});
