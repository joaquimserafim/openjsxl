import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { openXlsx, streamSheetRows } from "../workbook"

// Missing optional parts must degrade gracefully, never throw (F2.4b). minimal.xlsx is a
// generated workbook with numbers and a boolean only, so it carries neither sharedStrings.xml
// nor styles.xml.

describe("reader — missing optional parts", () => {
	it("opens a workbook that has no styles.xml or sharedStrings.xml", async () => {
		const wb = await openXlsx(await loadFixture("minimal.xlsx"))
		expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1"])
		const sheet = wb.sheet("Sheet1")
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "number", value: 1 })
		expect(sheet.cell("B1")).toEqual({ ref: "B1", type: "boolean", value: true })
		// No style table ⇒ no resolvable number format, but no error.
		expect(sheet.numberFormat("A1")).toBeUndefined()
		// Metadata accessors are empty, not throwing.
		expect(sheet.mergedCells).toEqual([])
		expect(sheet.hyperlinks).toEqual([])
		expect(sheet.comments).toEqual([])
	})

	it("streams rows from a workbook with no optional parts", async () => {
		const bytes = await loadFixture("minimal.xlsx")
		const rows = []
		for await (const row of streamSheetRows(bytes)) rows.push(row.cells.length)
		expect(rows).toEqual([2])
	})
})

describe("reader — maxPartBytes (zip-bomb guard)", () => {
	it("rejects a part larger than the configured limit", async () => {
		const bytes = await loadFixture("basic.xlsx")
		await expect(openXlsx(bytes, { maxPartBytes: 1 })).rejects.toMatchObject({
			code: "part-too-large",
		})
	})

	it("opens normally when the limit is generous", async () => {
		const bytes = await loadFixture("basic.xlsx")
		const wb = await openXlsx(bytes, { maxPartBytes: 1_000_000 })
		expect(wb.sheet("Sheet1").cell("A1")).toEqual({ ref: "A1", type: "string", value: "hello" })
	})

	it("applies the limit to streaming too", async () => {
		const bytes = await loadFixture("basic.xlsx")
		await expect(
			streamSheetRows(bytes, undefined, { maxPartBytes: 1 }).next(),
		).rejects.toMatchObject({ code: "part-too-large" })
	})
})
