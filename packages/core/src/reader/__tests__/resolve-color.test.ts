import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { openXlsx } from "../workbook"

// F5.3 — Workbook.resolveColor against real producer output. The tint algorithm itself is unit-
// tested in ooxml/__tests__/theme.test.ts against Excel swatches; this pins the end-to-end path:
// a cell's raw {theme, tint} resolved against the workbook's actual theme part.

describe("Workbook.resolveColor — default-theme fixture (e2e)", () => {
	it("resolves a themed font color against the workbook theme", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-styled.xlsx"))
		const color = wb.sheet("Styled").style("B2")?.font?.color
		expect(color).toEqual({ theme: 4, tint: 0.3999755851924192 }) // accent1 (4F81BD) lighter 40%
		expect(wb.resolveColor(color as { theme: number; tint: number })).toBe("FF96B4D8")
	})

	it("passes an rgb color straight through as 8-digit ARGB", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-styled.xlsx"))
		const fill = wb.sheet("Styled").style("C5")?.fill?.fgColor
		expect(fill).toEqual({ rgb: "FFDDEBF7" })
		expect(wb.resolveColor(fill as { rgb: string })).toBe("FFDDEBF7")
	})
})

describe("Workbook.resolveColor — custom-theme fixture (e2e)", () => {
	it("resolves theme colors against the CUSTOM theme, not the default", async () => {
		// openpyxl-customtheme.xlsx recolors accent1 (theme index 4) to pure red FF0000.
		const wb = await openXlsx(await loadFixture("openpyxl-customtheme.xlsx"))
		const sheet = wb.sheet("Themed")
		expect(wb.resolveColor(sheet.style("B2")?.font?.color as { theme: number })).toBe(
			"FFFF0000",
		)
		expect(
			wb.resolveColor(sheet.style("C3")?.font?.color as { theme: number; tint: number }),
		).toBe("FFFF6666") // accent1 (FF0000) lighter 40%
	})

	it("exposes the raw theme part", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-customtheme.xlsx"))
		expect(wb.themeXml).toContain("FF0000") // the recolored accent1
	})
})

describe("Workbook.resolveColor — no theme part", () => {
	it("returns undefined for a theme color when the workbook has no theme", async () => {
		// basic.xlsx has no styles/theme part at all.
		const wb = await openXlsx(await loadFixture("basic.xlsx"))
		expect(wb.themeXml).toBeUndefined()
		expect(wb.resolveColor({ theme: 4 })).toBeUndefined()
		expect(wb.resolveColor({ rgb: "FF00FF00" })).toBe("FF00FF00") // rgb still resolves
	})
})
