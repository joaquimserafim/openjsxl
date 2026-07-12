import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";

// F8.4 — end-to-end integration + the oracle corpus. Three angles:
//   1. A realistic multi-function workbook evaluated whole; expected values are DOCUMENTED Excel,
//      cross-checked cell-for-cell against Python `formulas` out-of-tree (probes stay in the scratchpad).
//   2. Evaluation-vs-cache agreement on REAL producer fixtures — our re-evaluation of an openpyxl-
//      authored formula matches the `<v>` that producer cached, an independent cross-engine check.
//   3. The stale-cache contract (decision 2): evaluation is read-only and never trusts the cached
//      `<v>`, so an out-of-date cache is superseded by our computed value (the divergence is NAMED).

const f = (formula: string, value: number | string = 0): CellInput => ({ formula, value });

describe("formula integration — a realistic workbook evaluates end to end", () => {
	it("evaluates SUM/AVERAGE/IF/VLOOKUP/SUMIF/COUNTIF/IFERROR/nested against oracle-agreed values", async () => {
		// Data in A1:C4, a report down column E. Expected values were cross-checked against Python
		// `formulas` out-of-tree; here we pin the whole-workbook evaluation end to end.
		const wb = await writeXlsx({
			sheets: [
				{
					name: "Sales",
					rows: [
						[10, "apple", 1, null, f("SUM(A1:A4)")],
						[20, "banana", 2, null, f("AVERAGE(A1:A4)")],
						[30, "apple", 3, null, f('IF(E1>50,"big","small")')],
						[40, "cherry", 4, null, f('SUMIF(B1:B4,"apple",A1:A4)')],
						[null, null, null, null, f("VLOOKUP(30,A1:C4,3,FALSE)")],
						[null, null, null, null, f("ROUND(E2/3,2)")],
						[null, null, null, null, f('CONCAT(B1,"-",B2)')],
						[null, null, null, null, f('IFERROR(A1/0,"n/a")')],
						[null, null, null, null, f('COUNTIF(B1:B4,"apple")')],
						[null, null, null, null, f("E1+E4")],
						[null, null, null, null, f("MAX(A1:A4)-MIN(A1:A4)")],
						[null, null, null, null, f("IF(AND(A1>5,A2>5),SUM(C1:C4),0)")],
					],
				},
			],
		});
		const book = await openXlsx(wb);
		const r = await evaluateWorkbook(book);
		const g = (ref: string) => r.get("Sales", ref);
		expect(g("E1")).toBe(100); // SUM
		expect(g("E2")).toBe(25); // AVERAGE
		expect(g("E3")).toBe("big"); // nested IF over E1
		expect(g("E4")).toBe(40); // SUMIF apple → 10 + 30
		expect(g("E5")).toBe(3); // VLOOKUP 30 → C3
		expect(g("E6")).toBe(8.33); // ROUND(25/3, 2)
		expect(g("E7")).toBe("apple-banana"); // CONCAT
		expect(g("E8")).toBe("n/a"); // IFERROR wraps #DIV/0!
		expect(g("E9")).toBe(2); // COUNTIF apple
		expect(g("E10")).toBe(140); // E1 + E4 (formula-to-formula chain)
		expect(g("E11")).toBe(30); // MAX - MIN
		expect(g("E12")).toBe(10); // IF(AND(...), SUM(C1:C4))
	});

	it("evaluates across sheets and reports every formula cell keyed by sheet!ref", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Data", rows: [[7], [f("A1*6")]] },
					{ name: "Report", rows: [[f("Data!A2+8")]] },
				],
			}),
		);
		const r = await evaluateWorkbook(book);
		expect(r.get("Data", "A2")).toBe(42);
		expect(r.get("Report", "A1")).toBe(50); // 42 + 8, resolved cross-sheet
		expect(r.cells.some((c) => c.sheet === "Report" && c.ref === "A1")).toBe(true);
	});
});

describe("formula integration — evaluation agrees with a producer's cached values", () => {
	// The reader surfaces a formula cell's cached <v> as its value; re-evaluating must agree with what
	// the independent producer (openpyxl) computed. This is the cross-engine corpus check, in-tree.
	it("re-evaluates basic.xlsx's cached formula to the same value", async () => {
		const book = await openXlsx(await loadFixture("basic.xlsx"));
		const sheetName = book.sheets[0]?.name ?? "";
		const cached = book.sheet(sheetName).cell("E1").value; // openpyxl's cached B1*2
		expect(typeof cached).toBe("number");
		expect(await evaluateCell(book, sheetName, "E1")).toBe(cached);
	});

	it("re-evaluates shared-formula.xlsx's translated dependents to their cached values", async () => {
		const book = await openXlsx(await loadFixture("shared-formula.xlsx"));
		const sheetName = book.sheets[0]?.name ?? "";
		const ws = book.sheet(sheetName);
		// B1=A1*2 (master) and its shared-formula dependents B2/B3 — all carry openpyxl-cached <v>.
		for (const ref of ["B1", "B2", "B3"]) {
			const cached = ws.cell(ref).value;
			expect(typeof cached).toBe("number");
			expect(await evaluateCell(book, sheetName, ref)).toBe(cached);
		}
	});
});

describe("formula integration — 3-D references yield a typed error, never a silent value (decision 5)", () => {
	it("returns #REF! for Sheet1:Sheet3!A1 instead of silently binding to the first sheet", async () => {
		// The parser carries the 3-D span (SheetSpec.toName); the evaluator refuses it with a TYPED
		// error rather than collapsing to Sheet1's value (which would be silently wrong).
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Sheet1", rows: [[100]] },
					{ name: "Sheet2", rows: [[200]] },
					{ name: "Sheet3", rows: [[300]] },
					{ name: "Report", rows: [[f("SUM(Sheet1:Sheet3!A1)"), f("Sheet1:Sheet3!A1")]] },
				],
			}),
		);
		const r = await evaluateWorkbook(book);
		expect(r.get("Report", "A1")).toEqual(errorValue("#REF!")); // SUM over a 3-D range
		expect(r.get("Report", "B1")).toEqual(errorValue("#REF!")); // a bare 3-D reference
	});
});

describe("formula integration — evaluation supersedes a stale cache (decision 2)", () => {
	it("returns the computed value, not the out-of-date cached <v>", async () => {
		// Author a formula whose cached value is deliberately wrong; evaluation is read-only and never
		// trusts the cache, so the computed value wins. This is the NAMED evaluation-vs-cache divergence.
		const book = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[10], [{ formula: "A1+1", value: 999 }]] }],
			}),
		);
		expect(book.sheet("S").cell("A2").value).toBe(999); // the stale cache the reader surfaces
		expect(await evaluateCell(book, "S", "A2")).toBe(11); // our computed value supersedes it
	});
});
