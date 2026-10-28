import { describe, expect, it } from "vitest";
import {
	bestDoublingRatio,
	LINEAR_RATIO_LIMIT,
	MIN_SAMPLE_MS,
	threadCpuMs,
} from "../../ooxml/__tests__/complexity";
import { rewriteParts } from "../../reader/__tests__/shared-part-fixtures";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import type { CellInput } from "../../writer/types";
import { columnWalk, evaluateWorkbook, indexedPositions, lowerBound, sheetIndex } from "../eval";
import { errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.67: a range finds its populated cells through a per-sheet index of sorted rows (each with its
// sorted columns) and sorted columns (each with its sorted rows, D11). It walks whichever side visits
// fewer positions, so its cost follows its own cells, never the rest of the sheet.

// `bestDoublingRatio` (ooxml/__tests__/complexity.ts) for async work: each sample awaits one call and
// reads this thread's CPU time. Work on other threads (inflate) is not counted, which only lowers
// both sides; a sample shorter than MIN_SAMPLE_MS is retried with the call repeated.
async function asyncDoublingRatio(
	make: (size: number) => Promise<() => Promise<unknown>>,
	n: number,
): Promise<number> {
	const small = await make(n);
	const large = await make(2 * n);
	const time = async (fn: () => Promise<unknown>) => {
		for (let reps = 1; ; reps *= 2) {
			const start = threadCpuMs();
			for (let i = 0; i < reps; i++) await fn();
			const elapsed = threadCpuMs() - start;
			if (elapsed >= MIN_SAMPLE_MS) return elapsed / reps;
		}
	};
	let best = Number.POSITIVE_INFINITY;
	for (let attempt = 0; attempt < 3 && best >= LINEAR_RATIO_LIMIT; attempt++) {
		let fastestSmall = Number.POSITIVE_INFINITY;
		let fastestLarge = Number.POSITIVE_INFINITY;
		for (let s = 0; s < 3; s++) {
			fastestSmall = Math.min(fastestSmall, await time(small));
			fastestLarge = Math.min(fastestLarge, await time(large));
		}
		best = Math.min(best, fastestLarge / fastestSmall);
	}
	return best;
}

// A sheet model: row → (column → cell), filled in the order given.
function model(cells: readonly (readonly [number, number])[]): Map<number, Map<number, string>> {
	const rows = new Map<number, Map<number, string>>();
	for (const [col, row] of cells) {
		let cols = rows.get(row);
		if (cols === undefined) {
			cols = new Map();
			rows.set(row, cols);
		}
		cols.set(col, `${col},${row}`);
	}
	return rows;
}

// A call that evaluates the workbook `rows(n)` builds, after one warm-up evaluation.
function evaluation(rows: (n: number) => CellInput[][]) {
	return async (n: number) => {
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: rows(n) }] }));
		await evaluateWorkbook(book);
		return () => evaluateWorkbook(book);
	};
}

const formula = (text: string): CellInput => ({ formula: text, value: 0 });
// n rows of ten numbers in A:J.
const filled = (n: number, extra: (row: number) => CellInput[] = () => []) =>
	Array.from({ length: n }, (_, i): CellInput[] => [
		...Array.from({ length: 10 }, (_, c) => i + c),
		...extra(i + 1),
	]);

describe("a range over a long sheet", () => {
	it("costs time in proportion to the rows, for a moving average down every row", async () => {
		const rows = (n: number) =>
			Array.from({ length: n }, (_, i): CellInput[] => [
				i + 1,
				formula(`AVERAGE(A${i + 1}:A${i + 7})`),
			]);
		expect(await asyncDoublingRatio(evaluation(rows), 5_000)).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("costs time in proportion to the rows, for a count and a sum over an empty column on every row", async () => {
		const count = (n: number) =>
			Array.from({ length: n }, (_, i): CellInput[] => [i + 1, formula(`COUNT(C$1:C$${n})`)]);
		const sum = (n: number) =>
			Array.from({ length: n }, (_, i): CellInput[] => [i + 1, formula("SUM(D:D)")]);
		expect(await asyncDoublingRatio(evaluation(count), 5_000)).toBeLessThan(LINEAR_RATIO_LIMIT);
		expect(await asyncDoublingRatio(evaluation(sum), 5_000)).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("costs time in proportion to the rows, for a VLOOKUP into a 10-row side table on every row", async () => {
		// A:C hold the tall table and D1:E10 the side table; F looks each row up in D:E.
		const rows = (n: number) =>
			Array.from({ length: n }, (_, i): CellInput[] => [
				i + 1,
				(i + 1) * 2,
				(i + 1) * 3,
				i < 10 ? i + 1 : null,
				i < 10 ? (i + 1) * 10 : null,
				formula(`VLOOKUP(MOD(A${i + 1},10)+1,D:E,2,FALSE)`),
			]);
		expect(await asyncDoublingRatio(evaluation(rows), 5_000)).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("costs time in proportion to the cells, for one sum over a whole filled sheet", async () => {
		const sheetSum = (n: number) =>
			filled(n, (row) => (row === 1 ? [formula("SUM(A:J)")] : []));
		expect(await asyncDoublingRatio(evaluation(sheetSum), 5_000)).toBeLessThan(
			LINEAR_RATIO_LIMIT,
		);
	}, 120_000);

	it("costs time in proportion to the cells, for a totals row under a filled sheet", async () => {
		// Row n+1 sums each of the ten columns above it.
		const totalsRow = (n: number) => [
			...filled(n),
			Array.from({ length: 10 }, (_, c): CellInput => {
				const col = String.fromCharCode(65 + c);
				return formula(`SUM(${col}1:${col}${n})`);
			}),
		];
		expect(await asyncDoublingRatio(evaluation(totalsRow), 5_000)).toBeLessThan(
			LINEAR_RATIO_LIMIT,
		);
	}, 120_000);

	it("costs time in proportion to the rows, for row totals across a filled sheet", async () => {
		// HEAD scanned every row of the sheet for each total.
		const rowTotals = (n: number) => filled(n, (row) => [formula(`SUM(A${row}:J${row})`)]);
		expect(await asyncDoublingRatio(evaluation(rowTotals), 5_000)).toBeLessThan(
			LINEAR_RATIO_LIMIT,
		);
	}, 120_000);
});

describe("range results", () => {
	it("are what they were, for every shape of range", async () => {
		// Data holds 1..5 in A1:A5, 10 and 20 in C2 and E4, and 100 in B7.
		const rows: CellInput[][] = [
			[1],
			[2, null, 10],
			[3],
			[4, null, null, null, 20],
			[5],
			[],
			[null, 100],
		];
		expect(
			await evaluateFormulas(
				[
					"SUM(Data!A1:A5)",
					"SUM(Data!A5:A1)",
					"SUM(Data!A:A)",
					"SUM(Data!A:E)",
					"SUM(Data!2:4)",
					"SUM(Data!B2:D4)",
					"SUM(Data!C1:E7)",
					"COUNT(Data!F1:Z99)",
					"COUNT(Data!A8:E100)",
					"SUM(Data!A1:XFD1048576)",
					"COUNTA(Data!B:B)",
					"SUM(Data!4:4)",
				],
				rows,
			),
		).toEqual([15, 15, 15, 145, 39, 10, 30, 0, 0, 145, 1, 24]);
	});

	it("meet the first error row by row, also on a range walked by column", async () => {
		// Beside 100 rows in A, D5 is #DIV/0! and E3 is #N/A: row 3 comes first.
		const rows = Array.from({ length: 100 }, (_, i): CellInput[] => [
			i + 1,
			null,
			null,
			i === 4 ? formula("1/0") : null,
			i === 2 ? formula("NA()") : null,
		]);
		expect(await evaluateFormulas(["SUM(Data!D:E)", "SUM(Data!D1:E100)"], rows)).toEqual([
			errorValue("#N/A"),
			errorValue("#N/A"),
		]);
		const index = sheetIndex(
			model([
				...Array.from({ length: 100 }, (_, i): readonly [number, number] => [1, i + 1]),
				[4, 5],
				[5, 3],
			]),
		);
		expect(columnWalk(index, 4, 1, 5, 1048576)).toEqual([
			[5, 3],
			[4, 5],
		]);
	});

	it("end at the whole row below a row bound that is not a whole number, on either walk", async () => {
		// Only column A is filled, so the whole-row ranges walk their one column.
		expect(
			await evaluateFormulas(
				["SUM(Data!2:3.5)", "COUNT(Data!2:3.5)"],
				[[1], [2], [3], [4], [5]],
			),
		).toEqual([5, 2]);
		const column = sheetIndex(
			model(Array.from({ length: 5 }, (_, i): readonly [number, number] => [1, i + 1])),
		);
		expect(columnWalk(column, 1, 2, 16384, 3.5)).toEqual([
			[1, 2],
			[1, 3],
		]);
	});

	it("see the cells of the range's own sheet", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "X",
						rows: [
							[1, { formula: "SUM(A:A)", value: 0 }],
							[2, { formula: "SUM(Y!A:A)", value: 0 }],
						],
					},
					{
						name: "Y",
						rows: [
							[10, { formula: "SUM(A:A)", value: 0 }],
							[null, { formula: "SUM(X!A:A)", value: 0 }],
							[],
							[],
							[20],
						],
					},
				],
			}),
		);
		const result = await evaluateWorkbook(book);
		expect(
			["X!B1", "X!B2", "Y!B1", "Y!B2"].map((at) => result.get(at.slice(0, 1), at.slice(2))),
		).toEqual([3, 30, 30, 3]);
	});

	it("are read row by row, left to right, also where the file lists rows out of order", async () => {
		// A1 is #DIV/0! and A2 is #N/A; the first error a sum meets wins. The edited file lists row 2
		// first: the range still meets A1 first, as the sheet lays the cells out.
		const base = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ formula: "1/0", value: 0 },
							{ formula: "SUM(A1:A2)", value: 0 },
						],
						[{ formula: "NA()", value: 0 }],
					],
				},
			],
		});
		const swapped = await rewriteParts(base, {
			"xl/worksheets/sheet1.xml": (xml) =>
				xml.replace(/(<row r="1".*?<\/row>)(<row r="2".*?<\/row>)/s, "$2$1"),
		});
		for (const bytes of [base, swapped]) {
			const result = await evaluateWorkbook(await openXlsx(bytes));
			expect(result.get("S", "B1")).toEqual(errorValue("#DIV/0!"));
		}
	});
});

describe("lowerBound", () => {
	it("finds the first position at or past the value", () => {
		expect(lowerBound([2, 4, 6], 4)).toBe(1);
		expect(lowerBound([2, 4, 6], 5)).toBe(2);
		expect(lowerBound([2, 4, 6], 1)).toBe(0);
		expect(lowerBound([2, 4, 6], 7)).toBe(3);
		expect(lowerBound([], 1)).toBe(0);
		expect(lowerBound([3, 3, 3], 3)).toBe(0);
	});
});

describe("sheetIndex", () => {
	it("sorts the rows, and each row's columns, whatever order the model holds them in", () => {
		expect(
			sheetIndex(
				model([
					[2, 3],
					[1, 3],
					[5, 1],
				]),
			),
		).toEqual({
			rows: [1, 3],
			colsOfRow: [[5], [1, 2]],
			cols: [1, 2, 5],
			rowsOfCol: [[3], [3], [1]],
		});
	});

	it("lists each column's rows in ascending order", () => {
		expect(
			sheetIndex(
				model([
					[1, 9],
					[1, 2],
					[3, 5],
					[1, 5],
				]),
			),
		).toEqual({
			rows: [2, 5, 9],
			colsOfRow: [[1], [1, 3], [1]],
			cols: [1, 3],
			rowsOfCol: [[2, 5, 9], [5]],
		});
	});

	it("leaves out a row that holds no cell", () => {
		const cells = model([[1, 2]]);
		cells.set(1, new Map());
		expect(sheetIndex(cells)).toEqual({
			rows: [2],
			colsOfRow: [[1]],
			cols: [1],
			rowsOfCol: [[2]],
		});
		expect(sheetIndex(new Map())).toEqual({ rows: [], colsOfRow: [], cols: [], rowsOfCol: [] });
	});
});

// Cells at A1, B3, E3 and A9.
const small = sheetIndex(
	model([
		[1, 1],
		[2, 3],
		[5, 3],
		[1, 9],
	]),
);

describe("columnWalk", () => {
	it("leaves a range that spans no more rows than columns to the row walk", () => {
		expect(columnWalk(small, 1, 3, 5, 3)).toBeUndefined();
		expect(columnWalk(small, 1, 1, 5, 9)).toBeUndefined();
		expect(columnWalk(small, 1, 4, 9, 8)).toBeUndefined();
	});

	it("counts only the rows inside the range", () => {
		// Ten rows hold A; row 5 also holds B and C. A1:C5's one row beats its three columns.
		const tall = sheetIndex(
			model([
				...Array.from({ length: 10 }, (_, i): readonly [number, number] => [1, i + 1]),
				[2, 5],
				[3, 5],
			]),
		);
		expect(columnWalk(tall, 1, 5, 3, 5)).toBeUndefined();
		expect([...indexedPositions(tall, 1, 5, 3, 5)]).toEqual([
			[1, 5],
			[2, 5],
			[3, 5],
		]);
	});

	it("leaves it to the row walk too when sorting its finds would cost more", () => {
		expect(columnWalk(small, 1, 1, 2, 9)).toBeUndefined();
	});

	it("walks fewer columns than rows, giving what it finds row by row and left to right", () => {
		expect(columnWalk(small, 5, 1, 5, 9)).toEqual([[5, 3]]);
		expect(columnWalk(small, 4, 1, 4, 9)).toEqual([]);
		expect(columnWalk(small, 2, 1, 2, 9)).toEqual([[2, 3]]);
	});

	it("walks one column whenever it spans more rows, with no sort to pay for", () => {
		// A1:A9 hold values and every row also holds B; A's nine finds beat its nine rows.
		const both = sheetIndex(
			model(
				Array.from({ length: 9 }, (_, i): (readonly [number, number])[] => [
					[1, i + 1],
					[2, i + 1],
				]).flat(),
			),
		);
		expect(columnWalk(both, 1, 1, 1, 9)).toEqual(
			Array.from({ length: 9 }, (_, i) => [1, i + 1]),
		);
		expect(columnWalk(both, 1, 1, 2, 9)).toBeUndefined();
	});
});

describe("indexedPositions", () => {
	const index = small;

	it("gives the positions inside the range, row by row and left to right", () => {
		expect([...indexedPositions(index, 1, 1, 5, 9)]).toEqual([
			[1, 1],
			[2, 3],
			[5, 3],
			[1, 9],
		]);
		expect([...indexedPositions(index, 2, 1, 5, 3)]).toEqual([
			[2, 3],
			[5, 3],
		]);
	});

	it("keeps both edges of the range and nothing past them", () => {
		expect([...indexedPositions(index, 5, 3, 5, 3)]).toEqual([[5, 3]]);
		expect([...indexedPositions(index, 2, 3, 4, 8)]).toEqual([[2, 3]]);
		expect([...indexedPositions(index, 3, 3, 4, 3)]).toEqual([]);
	});

	it("gives nothing when no row or no column falls in the range", () => {
		expect([...indexedPositions(index, 2, 1, 4, 1)]).toEqual([]);
		expect([...indexedPositions(index, 1, 4, 9, 8)]).toEqual([]);
		expect([...indexedPositions(index, 1, 10, 9, 99)]).toEqual([]);
		expect([...indexedPositions(sheetIndex(new Map()), 1, 1, 9, 9)]).toEqual([]);
	});

	it("gives a column walk's finds in the same order as the row walk", () => {
		expect([...indexedPositions(small, 5, 1, 5, 9)]).toEqual([[5, 3]]);
		expect([...indexedPositions(small, 2, 1, 2, 9)]).toEqual([[2, 3]]);
	});

	it("costs time in proportion to what it yields, for a whole column beside a tall one", () => {
		const ratio = bestDoublingRatio((n) => {
			const index = sheetIndex(
				model([
					...Array.from({ length: n }, (_, i): readonly [number, number] => [1, i + 1]),
					[4, 2],
				]),
			);
			return () => {
				let found = 0;
				for (let row = 1; row <= n; row++) {
					for (const _ of indexedPositions(index, 4, 1, 4, 1048576)) found++;
				}
				return found;
			};
		}, 20_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("costs time in proportion to what it yields, for moving windows down every row", () => {
		// A and B are filled on every row. A one-column window takes the column walk; a two-column
		// window's 14 finds take the row walk.
		const ratio = (c2: number) =>
			bestDoublingRatio((n) => {
				const index = sheetIndex(
					model(
						Array.from({ length: n }, (_, i): (readonly [number, number])[] => [
							[1, i + 1],
							[2, i + 1],
						]).flat(),
					),
				);
				expect(columnWalk(index, 1, 1, c2, 7) === undefined).toBe(c2 === 2);
				return () => {
					let found = 0;
					for (let row = 1; row <= n; row++) {
						for (const _ of indexedPositions(index, 1, row, c2, row + 6)) found++;
					}
					return found;
				};
			}, 20_000);
		expect(ratio(1)).toBeLessThan(LINEAR_RATIO_LIMIT);
		expect(ratio(2)).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("finds what a scan of the whole sheet finds, in row-major order", () => {
		// A seeded generator (MINSTD) keeps the case set fixed.
		let seed = 7;
		const next = (limit: number) => {
			seed = (seed * 48271) % 2147483647;
			return 1 + (seed % limit);
		};
		const walks = { byColumn: 0, byRow: 0 };
		for (let round = 0; round < 300; round++) {
			const cells = Array.from({ length: next(40) }, (): readonly [number, number] => [
				next(12),
				next(30),
			]);
			const cellMap = model(cells);
			const index = sheetIndex(cellMap);
			for (let q = 0; q < 10; q++) {
				const [ca, cb, ra, rb] = [next(14), next(14), next(32), next(32)];
				const [c1, c2, r1, r2] = [
					Math.min(ca, cb),
					Math.max(ca, cb),
					Math.min(ra, rb),
					Math.max(ra, rb),
				];
				const scan: (readonly [number, number])[] = [];
				for (const [row, cols] of cellMap) {
					for (const col of cols.keys()) {
						if (row >= r1 && row <= r2 && col >= c1 && col <= c2) scan.push([col, row]);
					}
				}
				scan.sort((a, b) => a[1] - b[1] || a[0] - b[0]);
				expect([...indexedPositions(index, c1, r1, c2, r2)]).toEqual(scan);
				if (columnWalk(index, c1, r1, c2, r2) === undefined) walks.byRow++;
				else walks.byColumn++;
			}
		}
		// Both walks were checked.
		expect(walks.byColumn).toBeGreaterThan(300);
		expect(walks.byRow).toBeGreaterThan(300);
	});
});
