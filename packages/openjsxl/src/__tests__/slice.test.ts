import { loadFixture } from "@openjsxl/fixtures"
import { type Cell, openXlsx } from "openjsxl"
import { describe, expect, it } from "vitest"

// End-to-end vertical slice: drive the whole reader through the PUBLIC `openjsxl` entry
// point (facade → @openjsxl/core), proving the surface a user installs actually composes —
// container → relationships → shared strings → typed cells.

describe("openjsxl vertical slice — basic.xlsx", () => {
	it("opens a real .xlsx and lists its sheets through the public API", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"))
		expect(wb.sheets).toEqual([
			{ name: "Sheet1", path: "xl/worksheets/sheet1.xml", visible: true, state: "visible" },
		])
	})

	it("reads string, number, and boolean cells with the right types", async () => {
		const sheet = (await openXlsx(await loadFixture("basic.xlsx"))).sheet("Sheet1")
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "string", value: "hello" })
		expect(sheet.cell("B1")).toEqual({ ref: "B1", type: "number", value: 42 })
		expect(sheet.cell("D1")).toEqual({ ref: "D1", type: "boolean", value: true })
	})

	it("turns a sheet into JSON records keyed by column letter", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"))
		const sheet = wb.sheet(wb.sheets[0]?.name ?? "")

		const records: Array<Record<string, Cell["value"]>> = []
		for await (const row of sheet.rows()) {
			const record: Record<string, Cell["value"]> = {}
			for (const cell of row.cells) {
				record[cell.ref.replace(/\d+$/, "")] = cell.value
			}
			records.push(record)
		}

		// C1 is a date-formatted serial, so it reads back as a real Date (F2.1).
		expect(records).toEqual([
			{ A: "hello", B: 42, C: new Date(Date.UTC(2020, 0, 1)), D: true, E: 84 },
			{ A: "world", B: 3.14159 },
		])
	})
})
