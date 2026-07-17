import { openXlsx, writeXlsx } from "openjsxl";
import { evaluateWorkbook, parseFormula } from "openjsxl/formula";
import { describe, expect, it } from "vitest";

// The opt-in `openjsxl/formula` entry point (M8) had no in-repo consumer — only example 12 exercised
// it, and examples don't run in CI, so a break in the facade → @openjsxl/core/formula chain (or in the
// package's `./formula` export map) would ship silently. This drives the subpath the way an installed
// user does: read a workbook through `openjsxl`, evaluate through `openjsxl/formula`.
describe("openjsxl/formula — public subpath consumer", () => {
	it("parses a formula through the installed entry point", () => {
		expect(parseFormula("1+2")).toBeDefined();
	});

	it("evaluates a formula end to end via the public facade chain", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[10, 20], [{ formula: "A1+B1" }]] }],
		});
		const result = await evaluateWorkbook(await openXlsx(bytes));
		expect(result.get("S", "A2")).toBe(30);
	});
});
