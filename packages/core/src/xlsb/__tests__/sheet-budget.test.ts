import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { parseXlsbStyles } from "../parts";
import { MAX_XLSB_CELLS, parseXlsbSheet, setWithinBudget } from "../sheet";
import { fullBlank, fullRk, repeatedCellsBin, rowRecord, shortBlank, shortRk } from "./sheet-bin";
import { stylesBin } from "./styles-bin";

// F11.59: style 0, the default every lookup falls back to, is not stored, so blank style-0 records
// cost nothing; the cells and stored styles of one sheet fail typed past a budget below the Map limit.

const budgetError = expect.objectContaining({
	name: "XlsxError",
	code: "part-too-large",
	message: expect.stringContaining("cell budget"),
});
const parse = (bytes: readonly number[] | Uint8Array, maxCells?: number) =>
	parseXlsbSheet(Uint8Array.from(bytes), [], undefined, false, maxCells);

describe("style 0 is not stored", () => {
	it("keeps only non-default styles, and number formats read the same", () => {
		const sheet = parse([
			...rowRecord(0),
			...shortBlank(),
			...shortBlank(0),
			...shortRk(7),
			...shortRk(8, 2),
			...fullBlank(4, 3),
		]);
		expect([...sheet.cellStyles]).toEqual([
			["D1", 2],
			["E1", 3],
		]);
		expect([...sheet.cells.keys()]).toEqual(["C1", "D1"]);
		const styles = parseXlsbStyles(stylesBin([[164, "0.0"]], [164, 14, 2, 10]));
		for (const ref of ["A1", "B1", "C1"]) {
			expect(styles.formatCode(sheet.cellStyles.get(ref))).toBe(styles.formatCode(0));
		}
		expect(styles.formatCode(0)).toBe("0.0");
	});

	it("still lets a style-0 record replace an earlier style at the same ref", () => {
		const sheet = parse([...rowRecord(0), ...fullRk(0, 1, 2), ...fullBlank(0, 0)]);
		expect([...sheet.cellStyles]).toEqual([]);
		expect(sheet.cells.get("A1")?.value).toBe(1);
	});

	it("keeps nothing for a sheet of style-0 blanks", () => {
		const sheet = parse(repeatedCellsBin(100, 1_000, shortBlank()));
		expect([sheet.cells.size, sheet.cellStyles.size]).toEqual([0, 0]);
	});
});

describe("the cell budget", () => {
	it("fails typed once the cells pass the budget", () => {
		expect(() => parse(repeatedCellsBin(1, 3, shortRk(1)), 2)).toThrow(budgetError);
		expect(() => parse(repeatedCellsBin(1, 3, shortRk(1)), 2)).toThrow(
			"more than 2 cells, the cell budget",
		);
	});

	it("fails typed once the styled cells pass the budget, blank or not", () => {
		expect(() => parse(repeatedCellsBin(1, 3, shortBlank(1)), 2)).toThrow(budgetError);
		expect(() => parse(repeatedCellsBin(1, 3, shortBlank(1)), 2)).toThrow(
			"more than 2 cells with a non-default style, the cell budget",
		);
		expect(() => parse(repeatedCellsBin(1, 3, shortRk(1, 1)), 2)).toThrow(budgetError);
	});

	it("can fill the budget exactly", () => {
		const sheet = parse(repeatedCellsBin(1, 3, shortRk(1, 1)), 3);
		expect([sheet.cells.size, sheet.cellStyles.size]).toEqual([3, 3]);
	});

	it("charges nothing for style-0 blanks, or for a ref written again", () => {
		expect(parse(repeatedCellsBin(10, 100, shortBlank()), 0).cells.size).toBe(0);
		// Columns past XFD clamp to XFD, so these 17 records keep one cell.
		const clamped = [...rowRecord(0)];
		for (let col = 16_384; col <= 16_400; col++) clamped.push(...fullRk(col, col, 1));
		const sheet = parse(clamped, 1);
		expect([...sheet.cells.values()]).toEqual([{ ref: "XFD1", type: "number", value: 16_400 }]);
	});

	it("sits below V8's 2^24-entry Map limit", () => {
		expect(MAX_XLSB_CELLS).toBeLessThan(2 ** 24);
	});
});

describe("setWithinBudget", () => {
	it("sets a new ref below the budget", () => {
		const map = new Map<string, number>();
		setWithinBudget(map, "A1", 1, 1, "cells");
		expect([...map]).toEqual([["A1", 1]]);
	});

	it("fails typed on a new ref at the budget, and replaces a held ref", () => {
		const map = new Map([["A1", 1]]);
		expect(() => setWithinBudget(map, "B1", 2, 1, "things")).toThrow("more than 1 things");
		setWithinBudget(map, "A1", 3, 1, "cells");
		expect([...map]).toEqual([["A1", 3]]);
	});
});

describe("cost", () => {
	it("reads a sheet of style-0 blanks in time linear in its records", () => {
		const ratio = bestDoublingRatio((n) => {
			const bin = repeatedCellsBin(n, 1_000, shortBlank());
			return () => parseXlsbSheet(bin, [], undefined, false);
		}, 200);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 60_000);
});
