import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openXlsx } from "../workbook";

// F10.5 — documented drops: a file carrying features openjsxl does not model (row/column outline
// grouping, sheet tab colors, document properties) must READ CLEAN, keep its cell data across a
// rewrite, and simply not surface those features — never a bare throw, never a silent value change.
// (The corpus property in bridge-styles.test.ts additionally pins the value round-trip for this file.)

describe("documented drops read clean and round-trip (F10.5)", () => {
	it("outline grouping / tab color / docProps are dropped; data survives", async () => {
		const before = await openXlsx(await loadFixture("openpyxl-dropped-features.xlsx"));
		const s = before.sheet("Grouped");
		// Cell values read normally.
		expect(s.cell("A1").value).toBe("row1");
		expect(s.cell("A5").value).toBe("row5");
		// Outline grouping is not modelled: rows carrying only an outlineLevel surface no rowProperties,
		// and the column keeps its width but no outline data (the model has no field for it).
		expect(s.rowProperties.size).toBe(0);
		for (const col of s.columns) {
			expect(col).not.toHaveProperty("outlineLevel");
		}

		// A rewrite does not throw, and the cell data survives (only the unmodelled features drop).
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		const s2 = after.sheet("Grouped");
		expect(s2.cell("A1").value).toBe("row1");
		expect(s2.cell("A5").value).toBe("row5");
	});
});
