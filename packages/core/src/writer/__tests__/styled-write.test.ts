import { describe, expect, it } from "vitest"
import { XlsxError } from "../../errors"
import { openXlsx } from "../../reader/workbook"
import { openZip } from "../../zip"
import { writeXlsx } from "../workbook"

// F4.2 — styled-cell write input + the style interner. The contract under test:
//   1. A written style re-reads through Worksheet.style(ref) as EXACTLY the input (one shared
//      model — the round-trip is structural).
//   2. Identical styles intern to one xf; the s-attribute space stays minimal.
//   3. Bare-value input keeps its pre-F4.2 bytes (golden-pinned styles.xml for the date case;
//      the full-archive byte-identity against the old writer was verified out-of-band).
//   4. Bad style input throws invalid-input naming the cell.

const decoder = new TextDecoder()

async function stylesPartOf(bytes: Uint8Array): Promise<string | undefined> {
	const zip = openZip(bytes)
	return zip.has("xl/styles.xml") ? decoder.decode(await zip.read("xl/styles.xml")) : undefined
}

describe("writeXlsx — styled cells round-trip through style(ref)", () => {
	it("carries font, fill, border, and alignment back verbatim", async () => {
		const style = {
			font: { name: "Arial", size: 14, bold: true, color: { rgb: "FFFF0000" } },
			fill: { patternType: "solid", fgColor: { rgb: "FFFFFF00" } },
			border: {
				top: { style: "thin" },
				bottom: { style: "double", color: { rgb: "FF0070C0" } },
			},
			alignment: { horizontal: "center", wrapText: true, indent: 2 },
		} as const
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[{ value: "x", style }]] }] }),
		)
		expect(wb.sheet("S").style("A1")).toEqual(style)
		expect(wb.sheet("S").cell("A1")).toMatchObject({ type: "string", value: "x" })
	})

	it("keeps theme and indexed colors raw, and emits the theme part only when needed", async () => {
		const themed = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[{ value: 1, style: { font: { color: { theme: 4, tint: 0.4 } } } }]],
				},
			],
		})
		const zip = openZip(themed)
		expect(zip.has("xl/theme/theme1.xml")).toBe(true)
		const wb = await openXlsx(themed)
		expect(wb.sheet("S").style("A1")?.font?.color).toEqual({ theme: 4, tint: 0.4 })

		const unthemed = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[{ value: 1, style: { font: { color: { indexed: 10 } } } }]],
				},
			],
		})
		expect(openZip(unthemed).has("xl/theme/theme1.xml")).toBe(false)
	})

	it("writes a styled BLANK cell (<c r s/>) that re-reads as an empty cell with a style", async () => {
		const style = { border: { top: { style: "medium" } } } as const
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [["a", { value: null, style }]] }],
			}),
		)
		const sheet = wb.sheet("S")
		expect(sheet.cell("B1").type).toBe("empty")
		expect(sheet.style("B1")).toEqual(style)
		// The styled blank occupies its ref: the dimension covers it.
		expect(sheet.dimension).toBe("A1:B1")
	})

	it("a styled Date keeps the date number format alongside the style", async () => {
		const date = new Date(Date.UTC(2022, 0, 15))
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[{ value: date, style: { font: { bold: true } } }]] }],
			}),
		)
		const cell = wb.sheet("S").cell("A1")
		expect(cell.type).toBe("date")
		expect((cell.value as Date).getTime()).toBe(date.getTime())
		expect(wb.sheet("S").style("A1")).toEqual({
			numberFormat: "mm-dd-yy",
			font: { bold: true },
		})
	})

	it("normalizes defaults away: {bold: false} etc. behaves exactly like no style", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ value: 1, style: { font: { bold: false } } },
							{ value: 2, style: { alignment: { indent: 0, wrapText: false } } },
						],
					],
				},
			],
		})
		// Nothing non-default was interned → no styles.xml at all, as if the input were bare.
		expect(await stylesPartOf(bytes)).toBeUndefined()
		const wb = await openXlsx(bytes)
		expect(wb.sheet("S").style("A1")).toBeUndefined()
	})
})

describe("writeXlsx — style interning", () => {
	it("identical styles share one xf whether the object is shared or inlined", async () => {
		const shared = { font: { bold: true } }
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ value: 1, style: shared },
							{ value: 2, style: { font: { bold: true } } }, // structurally equal, new object
							{ value: 3, style: { font: { italic: true } } }, // distinct
						],
					],
				},
			],
		})
		const styles = (await stylesPartOf(bytes)) as string
		// default xf + bold xf + italic xf = 3; the two bold cells collapsed into one.
		expect(styles).toContain('<cellXfs count="3">')
		expect(styles).toContain('<fonts count="3">') // default + bold + italic
	})

	it("a style equal to the workbook defaults interns to xf 0 (no s attribute)", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[{ value: 1, style: { font: { name: "Calibri", size: 11 } } }]],
				},
			],
		})
		// Calibri 11 IS font 0 — the style resolves to the default format entirely.
		expect(await stylesPartOf(bytes)).toBeUndefined()
	})
})

describe("writeXlsx — golden styles.xml (byte-compat with the pre-F4.2 writer)", () => {
	it("a bare Date produces the exact legacy stylesheet", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[new Date(Date.UTC(2020, 0, 1))]] }],
		})
		expect(await stylesPartOf(bytes)).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
				'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
				'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
				'<borders count="1"><border/></borders>' +
				'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
				'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
				'<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
				'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
		)
	})

	it("unstyled input emits no styles.xml and no theme part", async () => {
		const zip = openZip(await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1]] }] }))
		expect(zip.has("xl/styles.xml")).toBe(false)
		expect(zip.has("xl/theme/theme1.xml")).toBe(false)
	})
})

describe("writeXlsx — style validation (invalid-input naming the cell)", () => {
	async function failure(rows: readonly (readonly unknown[])[]): Promise<string> {
		const e = await writeXlsx({
			sheets: [{ name: "S", rows: rows as never }],
		}).then(
			() => undefined,
			(err) => err,
		)
		expect(e).toBeInstanceOf(XlsxError)
		expect((e as XlsxError).code).toBe("invalid-input")
		return (e as XlsxError).message
	}

	it("rejects a stray object without a value property", async () => {
		expect(await failure([[{ val: 1 }]])).toMatch(/cell A1: .*value/)
	})

	it("rejects unknown keys on the cell, the style, and its components", async () => {
		expect(await failure([[{ value: 1, styles: {} }]])).toMatch(
			/only "value", "style", and "formula"/,
		)
		expect(await failure([[{ value: 1, style: { fonts: {} } }]])).toMatch(
			/unknown property "fonts"/,
		)
		expect(await failure([[{ value: 1, style: { font: { weight: 700 } } }]])).toMatch(
			/unknown property "weight"/,
		)
	})

	it("rejects a nested object value", async () => {
		expect(await failure([[{ value: { value: 1 } }]])).toMatch(/cannot be an object/)
	})

	it("rejects bad enums and malformed colors", async () => {
		expect(await failure([[{ value: 1, style: { fill: { patternType: "plaid" } } }]])).toMatch(
			/patternType/,
		)
		expect(
			await failure([[{ value: 1, style: { border: { top: { style: "wavy" } } } }]]),
		).toMatch(/border line styles/)
		expect(await failure([[{ value: 1, style: { font: { color: { rgb: "red" } } } }]])).toMatch(
			/6- or 8-digit hex/,
		)
		expect(
			await failure([[{ value: 1, style: { font: { underline: "singleAccounting" } } }]]),
		).toMatch(/underline/)
		expect(
			await failure([[{ value: 1, style: { alignment: { textRotation: 255 } } }]]),
		).toMatch(/textRotation/)
	})

	it('rejects colors on a "none" fill', async () => {
		expect(
			await failure([
				[
					{
						value: 1,
						style: { fill: { patternType: "none", fgColor: { rgb: "FFFF0000" } } },
					},
				],
			]),
		).toMatch(/cannot carry colors/)
	})

	it("names the failing cell in the error", async () => {
		expect(
			await failure([
				["ok", "fine"],
				["x", { value: 1, style: { font: 3 } }],
			]),
		).toMatch(/^cell B2:/)
	})
})

describe("writeXlsx — hostile style input (adversarial-review regressions)", () => {
	async function code(fn: () => Promise<unknown>): Promise<string | undefined> {
		const e = await fn().then(
			() => undefined,
			(err) => err,
		)
		expect(e).toBeInstanceOf(XlsxError)
		return (e as XlsxError).code
	}

	it("rejects a font name with a control character or lone surrogate", async () => {
		const bad = `Bad${String.fromCharCode(1)}Font`
		expect(
			await code(() =>
				writeXlsx({
					sheets: [{ name: "S", rows: [[{ value: 1, style: { font: { name: bad } } }]] }],
				}),
			),
		).toBe("invalid-input")
		const lone = `A${String.fromCharCode(0xd800)}B`
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						{ name: "S", rows: [[{ value: 1, style: { font: { name: lone } } }]] },
					],
				}),
			),
		).toBe("invalid-input")
	})

	it("rejects theme/indexed color values past the unsignedInt range (no exponential notation)", async () => {
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						{
							name: "S",
							rows: [[{ value: 1, style: { font: { color: { theme: 1e21 } } } }]],
						},
					],
				}),
			),
		).toBe("invalid-input")
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						{
							name: "S",
							rows: [
								[{ value: 1, style: { font: { color: { indexed: 2 ** 32 } } } }],
							],
						},
					],
				}),
			),
		).toBe("invalid-input")
	})

	it("rejects exotic objects (Map/Date/class instances) as style components", async () => {
		// A Map has no own enumerable keys, so an own-keys check alone would pass it — and reading
		// .size through the prototype then fabricated a "font size 2". Must throw instead.
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						{
							name: "S",
							// biome-ignore lint/suspicious/noExplicitAny: hostile input on purpose
							rows: [[{ value: 1, style: { font: new Map([["x", 1]]) as any } }]],
						},
					],
				}),
			),
		).toBe("invalid-input")
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						// biome-ignore lint/suspicious/noExplicitAny: hostile input on purpose
						{ name: "S", rows: [[{ value: 1, style: { fill: new Date() as any } }]] },
					],
				}),
			),
		).toBe("invalid-input")
	})

	it("is immune to getter-backed properties changing value between check and emission", async () => {
		// A getter that passes validation reads and returns markup for the emission read would
		// inject raw XML into styles.xml. Single-read validators make the checked value the
		// emitted value — whatever the getter does afterwards is irrelevant.
		let reads = 0
		const color = {
			get rgb() {
				reads++
				return reads < 3 ? "FFAA0000" : '"/><injected/>'
			},
		}
		const bytes = await writeXlsx({
			sheets: [
				// biome-ignore lint/suspicious/noExplicitAny: hostile input on purpose
				{ name: "S", rows: [[{ value: 1, style: { font: { color: color as any } } }]] },
			],
		})
		const styles = new TextDecoder().decode(await openZip(bytes).read("xl/styles.xml"))
		expect(styles).toContain('rgb="FFAA0000"')
		expect(styles).not.toContain("injected")
	})
})
