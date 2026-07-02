import { describe, expect, it } from "vitest"
import { XlsxError } from "../../errors"
import { openXlsx } from "../../reader/workbook"
import { openZip } from "../../zip"
import { workbookToInput } from "../from-workbook"
import { writeXlsx } from "../workbook"

// F4.5 — geometry write + bridge. The written geometry must re-read through the reader's own
// accessors verbatim, sit in schema order (dimension → sheetViews → cols → sheetData), and carry
// across the bridge; a geometry-free sheet must emit the exact pre-F4.5 bytes (covered by the
// existing golden pins, re-asserted structurally here).

const decoder = new TextDecoder()

describe("writeXlsx — geometry round-trip", () => {
	it("writes columns, row properties, and a frozen pane that re-read verbatim", async () => {
		const columns = [
			{ min: 2, max: 3, width: 25.5 },
			{ min: 5, max: 5, hidden: true },
		] as const
		const rowProperties = { 2: { height: 33 }, 7: { hidden: true } } as const
		const freeze = { rows: 2, cols: 1 } as const
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "G", rows: [["a", "b"], [1]], columns, rowProperties, freeze }],
			}),
		)
		const sheet = wb.sheet("G")
		expect(sheet.columns).toEqual(columns)
		expect(Object.fromEntries(sheet.rowProperties)).toEqual(rowProperties)
		expect(sheet.freeze).toEqual(freeze)
		// Row 7 exists as a property-only row; its cells stay absent.
		expect(sheet.cell("A7").type).toBe("empty")
	})

	it("emits geometry in schema order: dimension, sheetViews, cols, sheetData", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "G",
					rows: [[1]],
					columns: [{ min: 1, max: 1, width: 10 }],
					freeze: { rows: 1 },
				},
			],
		})
		const xml = decoder.decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"))
		expect(xml).toContain(
			'<dimension ref="A1"/>' +
				'<sheetViews><sheetView workbookViewId="0">' +
				'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
				"</sheetView></sheetViews>" +
				'<cols><col min="1" max="1" width="10" customWidth="1"/></cols>' +
				"<sheetData>",
		)
	})

	it("merges row properties into rows that also have cells, ascending", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "G",
					rows: [["a"], undefined, ["c"]], // rows 1 and 3 have cells
					rowProperties: { 2: { height: 20 }, 3: { hidden: true } },
				},
			],
		})
		const xml = decoder.decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"))
		expect(xml).toContain(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>' +
				'<row r="2" ht="20" customHeight="1"/>' +
				'<row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>c</t></is></c></row>',
		)
	})

	it("normalizes no-op geometry away — the file matches a geometry-free write exactly", async () => {
		const plain = await writeXlsx({ sheets: [{ name: "G", rows: [[1]] }] })
		const noop = await writeXlsx({
			sheets: [
				{
					name: "G",
					rows: [[1]],
					columns: [{ min: 1, max: 1, hidden: false }],
					rowProperties: { 3: { hidden: false } },
					freeze: { rows: 0, cols: 0 },
				},
			],
		})
		expect(Array.from(noop)).toEqual(Array.from(plain))
	})

	it("carries geometry across the bridge", async () => {
		const first = await writeXlsx({
			sheets: [
				{
					name: "G",
					rows: [["x"]],
					columns: [{ min: 1, max: 2, width: 14.25 }],
					rowProperties: { 5: { height: 40.5, hidden: true } },
					freeze: { cols: 2 },
				},
			],
		})
		const again = await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(first))))
		const sheet = again.sheet("G")
		expect(sheet.columns).toEqual([{ min: 1, max: 2, width: 14.25 }])
		expect(Object.fromEntries(sheet.rowProperties)).toEqual({
			5: { height: 40.5, hidden: true },
		})
		expect(sheet.freeze).toEqual({ cols: 2 })
	})
})

describe("writeXlsx — geometry validation (invalid-input naming the sheet)", () => {
	async function failure(extra: object): Promise<string> {
		const e = await writeXlsx({
			sheets: [{ name: "G", rows: [[1]], ...extra } as never],
		}).then(
			() => undefined,
			(err) => err,
		)
		expect(e).toBeInstanceOf(XlsxError)
		expect((e as XlsxError).code).toBe("invalid-input")
		return (e as XlsxError).message
	}

	it("rejects bad column ranges, widths, and unknown keys", async () => {
		expect(await failure({ columns: [{ min: 0, max: 1, width: 10 }] })).toMatch(/columns\[0\]/)
		expect(await failure({ columns: [{ min: 2, max: 1, width: 10 }] })).toMatch(/min ≤ max/)
		expect(await failure({ columns: [{ min: 1, max: 1, width: 300 }] })).toMatch(/width/)
		expect(await failure({ columns: [{ min: 1, max: 1, width: 10, style: 2 }] })).toMatch(
			/unknown property "style"/,
		)
	})

	it("rejects bad row numbers and heights", async () => {
		expect(await failure({ rowProperties: { 0: { height: 10 } } })).toMatch(/row number/)
		expect(await failure({ rowProperties: { 2: { height: 500 } } })).toMatch(/height/)
		expect(await failure({ rowProperties: { 2: { ht: 10 } } })).toMatch(/unknown property "ht"/)
	})

	it("rejects bad freeze values", async () => {
		expect(await failure({ freeze: { rows: -1 } })).toMatch(/freeze\.rows/)
		expect(await failure({ freeze: { rows: 1.5 } })).toMatch(/freeze\.rows/)
		expect(await failure({ freeze: { panes: 1 } })).toMatch(/unknown property "panes"/)
	})
})
