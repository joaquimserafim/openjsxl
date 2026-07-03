import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";
import { parseMergedCells } from "../worksheet";

// Merged ranges (F2.3). `<mergeCells>` sits after `<sheetData>`; the accessor scans the token
// stream for `<mergeCell ref>` children and returns the A1 ranges in document order. Verified
// against real Excel output (merge_cells.xlsx, merged_range.xlsx — calamine, MIT).

describe("Worksheet.mergedCells — real fixtures", () => {
	it("reads the merged ranges of a single sheet in document order", async () => {
		const wb = await openXlsx(await loadFixture("merge_cells.xlsx"));
		expect(wb.sheet("Sheet1").mergedCells).toEqual(["A1:B1", "A2:A4", "B2:D4"]);
	});

	it("reads distinct ranges per sheet in a multi-sheet workbook", async () => {
		const wb = await openXlsx(await loadFixture("merged_range.xlsx"));
		expect(wb.sheet("Sheet1").mergedCells).toEqual([
			"H1:H2",
			"A1:A2",
			"B1:B2",
			"C1:D2",
			"C3:D3",
			"C4:D4",
			"E1:E2",
			"F1:F2",
			"G1:G2",
		]);
		expect(wb.sheet("Sheet2").mergedCells).toEqual([
			"A1:A4",
			"C3:D4",
			"F1:H4",
			"B1:B2",
			"C1:D2",
			"E1:E2",
		]);
	});

	it("is empty for a sheet with no merges", async () => {
		const wb = await openXlsx(await loadFixture("date.xlsx"));
		expect(wb.sheet("Sheet1").mergedCells).toEqual([]);
	});

	it("parses once and caches the result", async () => {
		const sheet = (await openXlsx(await loadFixture("merge_cells.xlsx"))).sheet("Sheet1");
		expect(sheet.mergedCells).toBe(sheet.mergedCells); // same reference ⇒ parsed lazily once
	});
});

describe("parseMergedCells — units", () => {
	it("collects refs from the mergeCells block", () => {
		const xml =
			'<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="C1:C3"/></mergeCells>';
		expect(parseMergedCells(xml)).toEqual(["A1:B2", "C1:C3"]);
	});

	it("tolerates a namespace prefix on the elements", () => {
		const xml = '<x:mergeCells><x:mergeCell ref="A1:B1"/></x:mergeCells>';
		expect(parseMergedCells(xml)).toEqual(["A1:B1"]);
	});

	it("skips a mergeCell with a missing or empty ref", () => {
		const xml =
			'<mergeCells><mergeCell/><mergeCell ref=""/><mergeCell ref="A1:A2"/></mergeCells>';
		expect(parseMergedCells(xml)).toEqual(["A1:A2"]);
	});

	it("returns empty when there is no mergeCells block", () => {
		expect(parseMergedCells('<sheetData><row r="1"/></sheetData>')).toEqual([]);
	});
});
