import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { FormulaError } from "../errors";
import { cellVisitBudget, evaluateCell, evaluateWorkbook } from "../eval";

// F11.41(a), owner call D17: `EvaluateOptions.maxCellVisits` is read once and checked before
// evaluation starts. It is a number, not NaN, at least 0, with Infinity for no limit; undefined and
// null keep the default of 5,000,000; anything else fails with FormulaError("unsupported"), so no
// value can switch the budget off.

const BAD: readonly unknown[] = [
	Number.NaN,
	"abc",
	"5",
	"",
	-1,
	-0.5,
	Number.NEGATIVE_INFINITY,
	true,
	{},
	[],
	10n,
];
const MESSAGE =
	"options.maxCellVisits must be a number of cell visits, at least 0 (Infinity for no limit)";

const unsupported = expect.objectContaining({ name: "FormulaError", code: "unsupported" });

// A1 holds 1 and A2:A100 each add 1 to the cell above, so evaluating A100 visits 99 formula cells.
async function chain() {
	const rows: CellInput[][] = [[1]];
	for (let i = 2; i <= 100; i++) rows.push([{ formula: `A${i - 1}+1`, value: 0 }]);
	return openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
}

async function noFormulas() {
	return openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [[1, 2]] }] }));
}

describe("cellVisitBudget", () => {
	it("is 5,000,000 when absent, and the value when it is a number at least 0", () => {
		expect(cellVisitBudget(undefined)).toBe(5_000_000);
		expect(cellVisitBudget(null)).toBe(5_000_000);
		for (const value of [0, 1, 2.5, 10_000_000, Number.POSITIVE_INFINITY]) {
			expect(cellVisitBudget(value)).toBe(value);
		}
	});

	it("refuses anything else with FormulaError unsupported", () => {
		for (const value of BAD) {
			expect(() => cellVisitBudget(value)).toThrow(FormulaError);
			expect(() => cellVisitBudget(value)).toThrow(unsupported);
			expect(() => cellVisitBudget(value)).toThrow(MESSAGE);
		}
	});
});

describe("evaluateWorkbook and evaluateCell check maxCellVisits (F11.41(a))", () => {
	it("refuse a bad budget before evaluating, even when no cell holds a formula", async () => {
		const books = [await chain(), await noFormulas()];
		for (const book of books) {
			for (const maxCellVisits of BAD) {
				const options = { maxCellVisits } as never;
				await expect(evaluateWorkbook(book, options)).rejects.toThrow(unsupported);
				await expect(evaluateCell(book, "S", "A1", options)).rejects.toThrow(unsupported);
			}
		}
	});

	it("refuse a bad budget before reading any sheet", async () => {
		const book = await chain();
		let sheetReads = 0;
		const counted = new Proxy(book, {
			get(target, key) {
				if (key === "sheet") {
					return (name: string) => {
						sheetReads++;
						return target.sheet(name);
					};
				}
				const value: unknown = Reflect.get(target, key, target);
				return typeof value === "function" ? value.bind(target) : value;
			},
		});
		await expect(evaluateWorkbook(counted, { maxCellVisits: Number.NaN })).rejects.toThrow(
			unsupported,
		);
		await expect(evaluateCell(counted, "S", "A100", { maxCellVisits: -1 })).rejects.toThrow(
			unsupported,
		);
		expect(sheetReads).toBe(0);
		expect((await evaluateWorkbook(counted)).get("S", "A100")).toBe(100);
		expect(sheetReads).toBeGreaterThan(0);
	});

	it("never let NaN switch the budget off", async () => {
		const book = await chain();
		await expect(
			evaluateCell(book, "S", "A100", { maxCellVisits: Number.NaN }),
		).rejects.toThrow(MESSAGE);
	});

	it("evaluate as with no options for Infinity, undefined and null", async () => {
		const book = await chain();
		const plain = await evaluateWorkbook(book);
		expect(plain.get("S", "A100")).toBe(100);
		for (const maxCellVisits of [Number.POSITIVE_INFINITY, undefined, null]) {
			const options = { maxCellVisits } as never;
			expect((await evaluateWorkbook(book, options)).cells).toEqual(plain.cells);
			expect(await evaluateCell(book, "S", "A100", options)).toBe(100);
		}
	});

	it("take 0 as a budget no formula fits, and a budget exactly the visits needed", async () => {
		const book = await chain();
		await expect(evaluateCell(book, "S", "A100", { maxCellVisits: 0 })).rejects.toMatchObject({
			code: "budget-exceeded",
		});
		expect((await evaluateWorkbook(await noFormulas(), { maxCellVisits: 0 })).cells).toEqual(
			[],
		);
		expect(await evaluateCell(book, "S", "A100", { maxCellVisits: 99 })).toBe(100);
		await expect(evaluateCell(book, "S", "A100", { maxCellVisits: 98 })).rejects.toMatchObject({
			code: "budget-exceeded",
		});
	});

	it("read the budget once per call", async () => {
		const book = await chain();
		let reads = 0;
		const options = {
			get maxCellVisits() {
				return reads++ === 0 ? 99 : Number.NaN;
			},
		};
		expect(await evaluateCell(book, "S", "A100", options)).toBe(100);
		expect(reads).toBe(1);
	});
});
