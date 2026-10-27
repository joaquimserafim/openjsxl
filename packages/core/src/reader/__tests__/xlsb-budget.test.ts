import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it, vi } from "vitest";
import { openXlsb } from "../xlsb";

// F11.59: openXlsb parses each sheet against MAX_XLSB_CELLS. The budget is lowered here so a test can
// reach it; the real one is 16,000,000.
vi.mock("../../xlsb", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../xlsb")>();
	return { ...actual, MAX_XLSB_CELLS: 2 };
});

describe("openXlsb's cell budget", () => {
	it("fails typed on a sheet with more than MAX_XLSB_CELLS cells", async () => {
		await expect(openXlsb(await loadFixture("xlsb-basic.xlsb"))).rejects.toEqual(
			expect.objectContaining({
				name: "XlsxError",
				code: "part-too-large",
				message: expect.stringContaining("more than 2 cells"),
			}),
		);
	});

	it("reads sheets within MAX_XLSB_CELLS", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-date1904.xlsb"));
		expect(wb.sheet("Sheet1").dimension).toBe("A1:B1");
	});
});
