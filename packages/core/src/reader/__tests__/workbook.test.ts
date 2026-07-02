import { loadFixture } from "@openjsxl/fixtures"
import { beforeAll, describe, expect, it } from "vitest"
import { openXlsx, streamSheetRows, type Workbook } from "../workbook"
import type { Row } from "../worksheet"

describe("openXlsx — basic.xlsx", () => {
	let wb: Workbook

	beforeAll(async () => {
		wb = await openXlsx(await loadFixture("basic.xlsx"))
	})

	it("lists the workbook sheets resolved through the relationship graph", () => {
		expect(wb.sheets).toEqual([
			{ name: "Sheet1", path: "xl/worksheets/sheet1.xml", visible: true },
		])
	})

	it("reads typed cells by A1 reference, including a date", () => {
		const sheet = wb.sheet("Sheet1")
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "string", value: "hello" })
		expect(sheet.cell("B1")).toEqual({ ref: "B1", type: "number", value: 42 })
		// C1 is serial 43831 with a date number format → a real Date now (F2.1).
		expect(sheet.cell("C1")).toEqual({
			ref: "C1",
			type: "date",
			value: new Date(Date.UTC(2020, 0, 1)),
		})
		expect(sheet.cell("D1")).toEqual({ ref: "D1", type: "boolean", value: true })
		expect(sheet.cell("E1")).toEqual({ ref: "E1", type: "number", value: 84 })
		expect(sheet.cell("A2")).toEqual({ ref: "A2", type: "string", value: "world" })
		expect(sheet.cell("B2")).toEqual({ ref: "B2", type: "number", value: 3.14159 })
	})

	it("returns an empty cell for an absent reference", () => {
		expect(wb.sheet("Sheet1").cell("Z99")).toEqual({ ref: "Z99", type: "empty", value: null })
	})

	it("iterates rows with for await", async () => {
		const seen: Array<{ index: number; refs: string[] }> = []
		for await (const row of wb.sheet("Sheet1").rows()) {
			seen.push({ index: row.index, refs: row.cells.map((c) => c.ref) })
		}
		expect(seen).toEqual([
			{ index: 1, refs: ["A1", "B1", "C1", "D1", "E1"] },
			{ index: 2, refs: ["A2", "B2"] },
		])
	})

	it("throws a helpful error for an unknown sheet name", () => {
		expect(() => wb.sheet("Nope")).toThrow(/no sheet named "Nope".*Sheet1/)
	})
})

describe("openXlsx — source types", () => {
	it("accepts an ArrayBuffer as well as a Uint8Array", async () => {
		const bytes = await loadFixture("basic.xlsx")
		const copy = bytes.slice() // a standalone ArrayBuffer-backed view
		const buffer = copy.buffer.slice(copy.byteOffset, copy.byteOffset + copy.byteLength)
		const wb = await openXlsx(buffer)
		expect(wb.sheet("Sheet1").cell("A1").value).toBe("hello")
	})

	it("accepts a Uint8Array that is a subarray view at a non-zero offset", async () => {
		const bytes = await loadFixture("basic.xlsx")
		const padded = new Uint8Array(bytes.length + 8)
		padded.set(bytes, 5)
		const view = padded.subarray(5, 5 + bytes.length) // byteOffset === 5
		const wb = await openXlsx(view)
		expect(wb.sheet("Sheet1").cell("A1").value).toBe("hello")
	})
})

describe("streamSheetRows — public streaming API", () => {
	async function collect(gen: AsyncIterable<Row>): Promise<Row[]> {
		const out: Row[] = []
		for await (const row of gen) out.push(row)
		return out
	}

	it("streams the default sheet identically to the eager reader (dates included)", async () => {
		const bytes = await loadFixture("basic.xlsx")
		const eager = await collect((await openXlsx(bytes)).sheet("Sheet1").rows())
		const streamed = await collect(streamSheetRows(bytes))
		expect(streamed).toEqual(eager)
		// And the date-styled C1 survives the streaming path as a Date.
		const c1 = streamed[0]?.cells.find((c) => c.ref === "C1")
		expect(c1).toEqual({ ref: "C1", type: "date", value: new Date(Date.UTC(2020, 0, 1)) })
	})

	it("selects a named sheet and throws on an unknown one", async () => {
		const bytes = await loadFixture("basic.xlsx")
		expect(await collect(streamSheetRows(bytes, "Sheet1"))).toHaveLength(2)
		await expect(collect(streamSheetRows(bytes, "Nope"))).rejects.toThrow(
			/no sheet named "Nope".*Sheet1/,
		)
	})
})
