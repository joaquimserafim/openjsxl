// 12 — Formulas: parse and evaluate (0.8): the opt-in engine behind openjsxl/formula.
//
//   node 12-formulas.mjs   (from ./examples)
//   pnpm --filter openjsxl-examples formulas
//
// The reader keeps a formula's TEXT and its cached value; `openjsxl/formula` adds an opt-in,
// zero-dependency ENGINE that recomputes them. It ships 90+ built-in functions, lets you register
// your own (UDFs), and is deterministic: volatile functions (TODAY/RAND) only work when you inject a
// clock/RNG, and circular references resolve to a dedicated #CYCLE! value instead of hanging. The
// engine never mutates the workbook — evaluation is read-only, so it can supersede a stale cache.

import { openXlsx, writeXlsx } from "openjsxl";
import { evaluateCell, evaluateWorkbook, FormulaError, parseFormula } from "openjsxl/formula";

const f = (formula) => ({ formula, value: 0 }); // a formula cell; the cached value is a placeholder
const show = (v) => (v && typeof v === "object" && v.kind === "error" ? v.code : v); // error → its code

// ── 1. Parse a formula string into a typed AST ────────────────────────────────
const ast = parseFormula("SUM(A1:C1) * 2");
console.log("parseFormula   :", ast.type, "→", JSON.stringify(ast));

// ── 2. Author a workbook of live formulas, then evaluate the whole thing ──────
const book = await openXlsx(
	await writeXlsx({
		sheets: [
			{
				name: "Sheet1",
				rows: [
					[10, 20, 30], // A1 B1 C1 — data
					[f("SUM(A1:C1)")], // A2 → 60
					[f('IF(A2>50,"over","under")')], // A3 → "over" (depends on A2)
					[f("A1/0")], // A4 → #DIV/0! (errors are values; they propagate)
					[f('IFERROR(A4,"recovered")')], // A5 → "recovered"
					[f("ROUND(A2/7,2)")], // A6 → 8.57
					[f("MILESTOKM(A1)")], // A7 → a user-defined function
				],
			},
		],
	}),
);

// A user-defined function, registered exactly the way the built-ins are.
const functions = {
	MILESTOKM: {
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => (typeof args[0] === "number" ? Math.round(args[0] * 160.934) / 100 : 0),
	},
};

const result = await evaluateWorkbook(book, { functions });
console.log("A2  SUM        :", show(result.get("Sheet1", "A2")));
console.log("A3  IF         :", show(result.get("Sheet1", "A3")));
console.log("A4  error      :", show(result.get("Sheet1", "A4")));
console.log("A5  IFERROR    :", show(result.get("Sheet1", "A5")));
console.log("A6  ROUND      :", show(result.get("Sheet1", "A6")));
console.log("A7  UDF        :", show(result.get("Sheet1", "A7")), "km");

// The same engine can resolve a single cell (and only its transitive dependencies).
console.log("evaluateCell A2:", show(await evaluateCell(book, "Sheet1", "A2", { functions })));

// ── 3. Circular references resolve to #CYCLE!, never a hang ────────────────────
const cyclic = await openXlsx(
	await writeXlsx({ sheets: [{ name: "S", rows: [[f("B1"), f("A1")]] }] }),
);
const cy = await evaluateWorkbook(cyclic);
console.log("cycle A1       :", show(cy.get("S", "A1")), "(B1 →", show(cy.get("S", "B1")), ")");

// ── 4. Volatile functions are gated on injected sources (determinism) ──────────
const vol = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [[f("TODAY()")]] }] }));
try {
	await evaluateCell(vol, "S", "A1"); // no clock injected → a typed refusal, not a nondeterministic value
} catch (err) {
	console.log("volatile gated :", err instanceof FormulaError ? err.code : String(err));
}
const today = await evaluateCell(vol, "S", "A1", { now: () => new Date(Date.UTC(2020, 0, 15)) });
console.log("TODAY injected :", today, "(the Excel date serial, deterministic)");
