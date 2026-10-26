import { describe, expect, it, vi } from "vitest";
import { openCsv } from "../csv";

// F11.58: openCsv places cells against MAX_CSV_CELLS. The budget is lowered here so a test can reach
// it; the real one is 16,000,000.
vi.mock("../../csv", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../csv")>();
	return { ...actual, MAX_CSV_CELLS: 3 };
});

describe("openCsv's cell budget", () => {
	it("reads up to MAX_CSV_CELLS non-empty cells", () => {
		const s = openCsv("a,,b\n\nc").sheet("Sheet1");
		expect([s.cell("A1").value, s.cell("C1").value, s.cell("A3").value]).toEqual([
			"a",
			"b",
			"c",
		]);
	});

	it("fails typed past MAX_CSV_CELLS", () => {
		expect(() => openCsv("a,b\nc,d")).toThrow(
			expect.objectContaining({
				name: "XlsxError",
				code: "part-too-large",
				message: expect.stringContaining("more than 3 non-empty cells"),
			}),
		);
	});
});
