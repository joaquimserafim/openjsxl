import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { parseStyles } from "../../ooxml"
import { openXlsx, Worksheet } from "../workbook"
import { parseCellStyles } from "../worksheet"

// Per-cell number formats (F2.3). Worksheet.numberFormat(ref) resolves a cell's style index
// through the style table to a format code — a custom code from <numFmts> or a built-in one.
// Verified against real output: date.xlsx (LibreOffice, custom codes) and basic.xlsx (our
// generator, built-in numFmtId 14).

describe("Worksheet.numberFormat — real fixtures", () => {
	it("reads custom number-format codes (LibreOffice date.xlsx)", async () => {
		const ws = (await openXlsx(await loadFixture("date.xlsx"))).sheet("Sheet1")
		expect(ws.numberFormat("A1")).toBe("yyyy\\-mm\\-dd") // custom numFmt 165
		expect(ws.numberFormat("A3")).toBe("[hh]:mm:ss") // custom numFmt 166 (elapsed time)
		expect(ws.numberFormat("B1")).toBe("General") // custom numFmt 164 = General
	})

	it("resolves a built-in numFmtId to its code (basic.xlsx C1 uses id 14)", async () => {
		const ws = (await openXlsx(await loadFixture("basic.xlsx"))).sheet("Sheet1")
		expect(ws.numberFormat("C1")).toBe("mm-dd-yy") // built-in 14
		expect(ws.numberFormat("A1")).toBe("General") // unstyled ⇒ default
	})

	it("is undefined when the workbook has no style table", () => {
		// A Worksheet built with a style-less context can't resolve a format.
		const ws = new Worksheet(
			{ name: "S", path: "x", visible: true, state: "visible" },
			'<c r="A1" s="1"><v>1</v></c>',
			{
				sharedStrings: [],
			},
		)
		expect(ws.numberFormat("A1")).toBeUndefined()
	})
})

describe("StyleTable.formatCode — units", () => {
	const styles = parseStyles(
		"<styleSheet>" +
			'<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>' +
			'<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="9"/><xf numFmtId="164"/></cellXfs>' +
			"</styleSheet>",
	)

	it("resolves built-in, custom, and default styles", () => {
		expect(styles.formatCode(0)).toBe("General") // built-in 0
		expect(styles.formatCode(1)).toBe("0%") // built-in 9
		expect(styles.formatCode(2)).toBe("0.00%") // custom 164
		expect(styles.formatCode(undefined)).toBe("General") // omitted s ⇒ style 0
	})

	it("is undefined for an out-of-range style index", () => {
		expect(styles.formatCode(99)).toBeUndefined()
	})
})

describe("parseCellStyles — units", () => {
	it("maps addressed cells to their style index, skipping cells without r or s", () => {
		const map = parseCellStyles(
			'<row><c r="A1" s="2"/><c r="B1"><v>1</v></c><c r="C1" s="0"/></row>',
		)
		expect(map.get("A1")).toBe(2)
		expect(map.has("B1")).toBe(false) // no s attribute
		expect(map.get("C1")).toBe(0)
	})
})
