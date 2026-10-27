import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it, vi } from "vitest";
import { openXlsb } from "../xlsb";

// F11.59: openXlsb reads the sheet list, shared strings and cell formats against
// MAX_XLSB_TABLE_ENTRIES. The budget is lowered to 3 here so the corpus can reach it; the real one is
// 16,000,000.
vi.mock("../../xlsb", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../xlsb")>();
	return { ...actual, MAX_XLSB_TABLE_ENTRIES: 3 };
});

const tableError = (what: string) =>
	expect.objectContaining({
		name: "XlsxError",
		code: "part-too-large",
		message: expect.stringContaining(`more than 3 ${what}`),
	});

describe("openXlsb's table budget", () => {
	it("fails typed on a workbook of 9 sheets", async () => {
		await expect(openXlsb(await loadFixture("xlsb-dimensions.xlsb"))).rejects.toEqual(
			tableError("sheets"),
		);
	});

	it("fails typed on 5 shared strings", async () => {
		await expect(openXlsb(await loadFixture("equiv.xlsb"))).rejects.toEqual(
			tableError("shared strings"),
		);
	});

	it("fails typed on 5 cell formats", async () => {
		await expect(
			openXlsb(await loadFixture("sheetjs-short-cells-inline.xlsb")),
		).rejects.toEqual(tableError("cell formats"));
	});

	it("reads a workbook at the budget: 3 sheets, 2 shared strings, 2 cell formats", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-basic.xlsb"));
		expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet2", "Hidden"]);
	});
});
