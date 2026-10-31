import { describe, expect, it } from "vitest";
import { evaluateWorkbook } from "../../formula/eval";
import { workbookToInput } from "../../writer/from-workbook";
import type { CellInput, SheetInput } from "../../writer/types";
import { writeXlsx } from "../../writer/workbook";
import { writeZip } from "../../writer/zip";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

// F11.9 reopening: reference-shaped text inside a name or a sheet qualifier must survive shared
// expansion. These are complete XLSX packages with real defined names, supporting sheets, and table
// metadata. Only the formula grouping is patched because the writer emits ordinary formulas.
async function sharedWorkbook(
	master: string,
	supportingSheets: readonly SheetInput[] = [],
	[first, second]: readonly [readonly CellInput[], readonly CellInput[]] = [
		[1, null],
		[2, null],
	],
) {
	const bytes = await writeXlsx({
		sheets: [
			{
				name: "S",
				rows: [
					[...first, { formula: master, value: 8 }],
					[...second, { formula: "0", value: 9 }],
				],
			},
			...supportingSheets,
		],
		definedNames: [{ name: "a1é", refersTo: "7" }],
	});
	const zip = openZip(bytes);
	const sheetPath = "xl/worksheets/sheet1.xml";
	const original = new TextDecoder().decode(await zip.read(sheetPath));
	const shared = original
		.replace("<f>", '<f t="shared" ref="C1:C2" si="0">')
		.replace("<f>0</f>", '<f t="shared" si="0"/>');
	expect(shared).toContain('<f t="shared" ref="C1:C2" si="0">');
	expect(shared).toContain('<f t="shared" si="0"/>');
	const parts: { name: string; data: Uint8Array }[] = [];
	for (const [name] of zip.entries) {
		parts.push({
			name,
			data: name === sheetPath ? new TextEncoder().encode(shared) : await zip.read(name),
		});
	}
	return openXlsx(await writeZip(parts), { verifyChecksums: true });
}

async function assertFormulaRoundTrip(book: Workbook, master: string, dependent: string) {
	expect(book.sheet("S").formula("C1")).toBe(master);
	expect(book.sheet("S").formula("C2")).toBe(dependent);
	const rewritten = await openXlsx(await writeXlsx(await workbookToInput(book)), {
		verifyChecksums: true,
	});
	expect(rewritten.sheet("S").formula("C1")).toBe(master);
	expect(rewritten.sheet("S").formula("C2")).toBe(dependent);
	return rewritten;
}

describe("shared formulas preserve identifiers and sheet qualifiers (F11.9)", () => {
	it("keeps a Unicode-suffixed name bound to its constant through expansion and the bridge", async () => {
		const book = await sharedWorkbook("a1é+A1");
		const result = await evaluateWorkbook(book);
		expect(result.get("S", "C1")).toBe(8);
		expect(result.get("S", "C2")).toBe(9);

		const rewritten = await assertFormulaRoundTrip(book, "a1é+A1", "a1é+A2");
		expect(rewritten.definedNames).toEqual([{ name: "a1é", refersTo: "7" }]);
		const evaluatedAgain = await evaluateWorkbook(rewritten);
		expect(evaluatedAgain.get("S", "C1")).toBe(8);
		expect(evaluatedAgain.get("S", "C2")).toBe(9);
	});

	it("preserves a Unicode-suffixed sheet name while shifting its cell reference", async () => {
		const book = await sharedWorkbook("a1é!A1+A1", [{ name: "a1é", rows: [[3], [4]] }]);
		const rewritten = await assertFormulaRoundTrip(book, "a1é!A1+A1", "a1é!A2+A2");
		expect(rewritten.sheet("a1é").cell("A2")).toEqual({
			ref: "A2",
			type: "number",
			value: 4,
		});
	});

	it("preserves a Unicode-suffixed table name and its structured reference", async () => {
		const book = await sharedWorkbook("SUM(b2é[X])+A1", [
			{
				name: "Data",
				rows: [["X"], [3], [4]],
				tables: [
					{
						name: "b2é",
						ref: "A1:A3",
						columns: [],
						headerRow: true,
						totalsRow: false,
					},
				],
			},
		]);
		const rewritten = await assertFormulaRoundTrip(book, "SUM(b2é[X])+A1", "SUM(b2é[X])+A2");
		expect(rewritten.sheet("Data").tables[0]?.name).toBe("b2é");
		expect(rewritten.sheet("Data").tables[0]?.columns).toEqual([{ name: "X" }]);
	});

	it("shifts both endpoints of an area with repeated sheet qualifiers", async () => {
		const book = await sharedWorkbook("SUM(Data!A1:Data!A2)", [
			{ name: "Data", rows: [[3], [4], [5]] },
		]);
		await assertFormulaRoundTrip(book, "SUM(Data!A1:Data!A2)", "SUM(Data!A2:Data!A3)");
	});

	it("shifts the surviving endpoint of an area ending in a reference error", async () => {
		const book = await sharedWorkbook("SUM(A1:#REF!)");
		await assertFormulaRoundTrip(book, "SUM(A1:#REF!)", "SUM(A2:#REF!)");
	});

	// Unquoted, `a1:b2!A1` is a range, as `parseFormula` reads it (F11.36, pinned below).
	it.each([
		["1:2", "1", "2"],
		["'a1:b2'", "a1", "b2"],
		["A1B:XFE1", "A1B", "XFE1"],
	])(
		"preserves the 3-D sheet span %s while shifting only its cell reference",
		async (span, from, to) => {
			const master = `SUM(${span}!A1)+A1`;
			const dependent = `SUM(${span}!A2)+A2`;
			const book = await sharedWorkbook(master, [
				{ name: from, rows: [[3], [4]] },
				{ name: to, rows: [[5], [6]] },
			]);
			// The evaluator deliberately does not support 3-D references; this pins reader/bridge fidelity.
			await assertFormulaRoundTrip(book, master, dependent);
		},
	);

	// F11.36: the unquoted `A1:` is a cell, as when the sheet is quoted. `A2:S!B2` is one range on S,
	// the sheet its qualified end names (`rangeSheet`, F11.65): 10 + 12 = 22. Reading `A1:S!` as a
	// 3-D span kept A1:B2, 33.
	it.each(["S", "'S'"])(
		"shifts the cell endpoint of a range ending in %s!B1 through expansion, evaluation and the bridge",
		async (sheet) => {
			const master = `SUM(A1:${sheet}!B1)`;
			const dependent = `SUM(A2:${sheet}!B2)`;
			const book = await sharedWorkbook(
				master,
				[],
				[
					[11, null],
					[10, 12],
				],
			);
			const result = await evaluateWorkbook(book);
			expect(result.get("S", "C1")).toBe(11);
			expect(result.get("S", "C2")).toBe(22);

			const rewritten = await assertFormulaRoundTrip(book, master, dependent);
			const evaluatedAgain = await evaluateWorkbook(rewritten);
			expect(evaluatedAgain.get("S", "C1")).toBe(11);
			expect(evaluatedAgain.get("S", "C2")).toBe(22);
		},
	);
});
