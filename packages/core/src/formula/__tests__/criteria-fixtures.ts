import { expect } from "vitest";
import type { DefinedName } from "../../ooxml/workbook";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateWorkbook } from "../eval";

type Rows = readonly (readonly CellInput[])[];

/**
 * Write `rows` to a `Data` sheet and each formula to `Results!A{n}`, evaluate the workbook, and return
 * the results in order. Results live on another sheet so whole-column ranges see only the data.
 */
export async function evaluateFormulas(
	formulas: readonly string[],
	rows: Rows,
	definedNames: readonly DefinedName[] = [],
) {
	const book = await openXlsx(
		await writeXlsx({
			sheets: [
				{ name: "Data", rows },
				{ name: "Results", rows: formulas.map((formula) => [{ formula, value: 0 }]) },
			],
			...(definedNames.length > 0 ? { definedNames } : {}),
		}),
	);
	const result = await evaluateWorkbook(book);
	return formulas.map((_, index) => result.get("Results", `A${index + 1}`));
}

/** The same single-criterion count written as COUNTIF and as COUNTIFS. */
export function countForms(range: string, criterion: string): readonly [string, string] {
	return [`COUNTIF(${range},${criterion})`, `COUNTIFS(${range},${criterion})`];
}

export type CountCase = readonly [range: string, criterion: string, expected: number];

/** Assert every case's COUNTIF and COUNTIFS forms both return its expected count. */
export async function expectCounts(
	cases: readonly CountCase[],
	rows: Rows,
	definedNames: readonly DefinedName[] = [],
): Promise<void> {
	const formulas = cases.flatMap(([range, criterion]) => countForms(range, criterion));
	const values = await evaluateFormulas(formulas, rows, definedNames);
	expect(values).toEqual(cases.flatMap(([, , expected]) => [expected, expected]));
}

/**
 * A header row plus 100 data rows (sheet rows 2-101). Row i has name `N{i}` in A (blank when i is a
 * multiple of 10), region North/East/South in B (i mod 3 = 1/2/0), amount i in C and numeric SKU
 * 1000 + i in D.
 */
export function salesTableRows(): Rows {
	const rows: CellInput[][] = [["Name", "Region", "Amount", "SKU"]];
	for (let i = 1; i <= 100; i++) {
		const region = i % 3 === 1 ? "North" : i % 3 === 2 ? "East" : "South";
		rows.push([i % 10 === 0 ? null : `N${i}`, region, i, 1000 + i]);
	}
	return rows;
}

/** Column A1:A9: text, number, boolean, blank, formula `""`, number, text, text, single-char text. */
export function mixedColumnRows(): Rows {
	return [
		["apple"],
		[3],
		[true],
		[null],
		[{ formula: '""', value: "" }],
		[2024],
		["2x"],
		["abcd"],
		["q"],
	];
}

/** A1:A3 = 5, TRUE, "c" with amounts 10, 20, 30 in B1:B3. */
export function relationalRows(): Rows {
	return [
		[5, 10],
		[true, 20],
		["c", 30],
	];
}

/**
 * A1:A5 and B1:B5 = 0, 0, blank, 5, "t"; amounts 1, 2, 4, 8, 16 in C1:C5 (distinct powers of two, so
 * every subset sums differently); column E stays empty.
 */
export function emptyCriterionRows(): Rows {
	return [
		[0, 0, 1],
		[0, 0, 2],
		[null, null, 4],
		[5, 5, 8],
		["t", "t", 16],
	];
}
