import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";

// Real-producer smoke tests. Our generator emits *stored* (uncompressed) entries with no
// namespace prefixes; genuine apps emit deflate-compressed, namespace-prefixed parts with full
// docProps. These fixtures are unmodified output from Microsoft Excel, LibreOffice, and
// openpyxl (vendored from calamine, MIT — see packages/fixtures/THIRD_PARTY.md), so a reader
// that only ever saw basic.xlsx can't pass by sharing its assumptions.

// One representative, non-empty cell per file: enough to prove the whole pipeline
// (deflate → tokenize → rels → shared strings → typed cell) ran end-to-end on real bytes.
const probes: Array<{ file: string; sheet: string; ref: string; type: string; value: unknown }> = [
	{ file: "merge_cells.xlsx", sheet: "Sheet1", ref: "A1", type: "string", value: "Row Merge" },
	{ file: "merged_range.xlsx", sheet: "Sheet1", ref: "A1", type: "string", value: "A1" },
	{ file: "hyperlinks.xlsx", sheet: "Links", ref: "A1", type: "string", value: "calamine repo" },
	{ file: "inventory-table.xlsx", sheet: "Sheet1", ref: "A1", type: "string", value: "Item" },
	{ file: "errors.xlsx", sheet: "Feuil1", ref: "A1", type: "error", value: "#DIV/0!" },
	{
		file: "date.xlsx",
		sheet: "Sheet1",
		ref: "A1",
		type: "date",
		value: new Date(Date.UTC(2021, 0, 1)),
	},
];

describe("real-producer fixtures — smoke", () => {
	for (const { file, sheet, ref, type, value } of probes) {
		it(`opens ${file} and reads ${ref} as ${type}`, async () => {
			const wb = await openXlsx(await loadFixture(file));
			const cell = wb.sheet(sheet).cell(ref);
			expect(cell.type).toBe(type);
			expect(cell.value).toEqual(value);
		});
	}

	it("applies the date1904 workbook flag (same serial, different anchor)", async () => {
		// Cell A3 holds the same time serial in both files; the 1904 system shifts it ~4 years.
		const v1900 = (await openXlsx(await loadFixture("date.xlsx"))).sheet("Sheet1").cell("A3");
		const v1904 = (await openXlsx(await loadFixture("date_1904.xlsx")))
			.sheet("Sheet1")
			.cell("A3");
		expect(v1900.type).toBe("date");
		expect(v1904.type).toBe("date");
		expect(v1900.value).toEqual(new Date(Date.UTC(1900, 0, 9, 15, 10, 10)));
		expect(v1904.value).toEqual(new Date(Date.UTC(1904, 0, 11, 15, 10, 10)));
	});

	it("reports sheet visibility from a real multi-state workbook", async () => {
		// any_sheets.xlsx has visible, hidden, and veryHidden sheets — both hidden states
		// collapse to `visible: false`.
		const wb = await openXlsx(await loadFixture("any_sheets.xlsx"));
		const byName = Object.fromEntries(wb.sheets.map((s) => [s.name, s.visible]));
		expect(byName).toMatchObject({
			Visible: true,
			Hidden: false,
			VeryHidden: false,
		});
	});
});
