import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { openXlsx } from "../workbook"

// F4.6 — structural metadata read against a real producer (openpyxl 3.1.5 — see
// fixtures/data/README.md): sheet visibility states, merged cells, and hyperlinks, read through
// the same accessors the bridge carries across. mergedCells/hyperlinks parsing itself is
// unit-tested since F2.x; this pins the state model end to end.

describe("sheet visibility — openpyxl-authored fixture (e2e)", () => {
	it("reads visible, hidden, and veryHidden states, with visible derived", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-metadata.xlsx"))
		expect(wb.sheets.map((s) => [s.name, s.state, s.visible])).toEqual([
			["Meta", "visible", true],
			["Hidden", "hidden", false],
			["Very", "veryHidden", false],
			["Plain", "visible", true],
		])
		expect(wb.sheet("Very").state).toBe("veryHidden")
		expect(wb.sheet("Meta").state).toBe("visible")
	})

	it("reads the fixture's merges and hyperlinks alongside the states", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-metadata.xlsx"))
		const meta = wb.sheet("Meta")
		expect(meta.mergedCells).toEqual(["A1:B2", "D1:D3"])
		expect(meta.hyperlinks).toEqual([
			{ ref: "A4", target: "https://example.com/docs", tooltip: "open the docs" },
			{ ref: "B4", location: "'Plain'!A1" },
		])
	})
})
