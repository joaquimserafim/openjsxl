import { readdir } from "node:fs/promises"
import { fileURLToPath } from "node:url"
import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { XlsxError } from "../../errors"
import { openXlsx, type Workbook } from "../../reader/workbook"
import { workbookToInput } from "../from-workbook"
import { writeXlsx } from "../workbook"

// F4.4 — the bridge carries styles. Contract: read → workbookToInput → writeXlsx → read gives a
// deep-equal style(ref) for every populated cell (values/types were already lossless since F3.3),
// and an UNSTYLED workbook still rewrites to byte-identical archives. The openpyxl-authored
// fixture is the acid test: real-producer styles, theme+tint colors, custom number formats.

async function styleSnapshot(wb: Workbook) {
	const out: Record<string, Record<string, unknown>> = {}
	for (const info of wb.sheets) {
		const sheet = wb.sheet(info.name)
		const cells: Record<string, unknown> = {}
		for await (const row of sheet.rows()) {
			for (const cell of row.cells) {
				cells[cell.ref] = {
					// Error cells write as their literal text (documented F3.3 flattening), so the
					// comparable identity of an 'error' cell IS the string it becomes.
					type: cell.type === "error" ? "string" : cell.type,
					value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
					style: sheet.style(cell.ref),
				}
			}
		}
		out[info.name] = {
			cells,
			// Geometry (F4.5) is part of the fidelity contract too.
			columns: sheet.columns,
			rowProperties: Object.fromEntries(sheet.rowProperties),
			freeze: sheet.freeze,
			// Structural metadata (F4.6): merges, hyperlinks, and tab visibility carry across.
			mergedCells: sheet.mergedCells,
			hyperlinks: sheet.hyperlinks,
			state: info.state,
		}
	}
	return out
}

async function rewrite(wb: Workbook): Promise<Uint8Array> {
	return writeXlsx(await workbookToInput(wb))
}

describe("bridge — styles round-trip", () => {
	it("carries every style of the openpyxl-authored fixture (acid test)", async () => {
		const before = await openXlsx(await loadFixture("openpyxl-styled.xlsx"))
		const snap = await styleSnapshot(before)
		const after = await openXlsx(await rewrite(before))
		expect(await styleSnapshot(after)).toEqual(snap)

		// Spot-check the hard cases survived: theme+tint, custom numFmt, full-load cell.
		const sheet = after.sheet("Styled")
		expect(sheet.style("B2")?.font?.color).toEqual({ theme: 4, tint: 0.3999755851924192 })
		expect(sheet.style("C4")?.numberFormat).toBe('"kg" 0.0')
		expect(sheet.style("C5")?.fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FFDDEBF7" },
		})
	})

	it("carries a styled BLANK cell across the bridge", async () => {
		const style = {
			border: { top: { style: "medium" } },
			fill: { patternType: "gray125" },
		} as const
		const first = await writeXlsx({
			sheets: [{ name: "S", rows: [["a", { value: null, style }]] }],
		})
		const again = await openXlsx(await rewrite(await openXlsx(first)))
		expect(again.sheet("S").cell("B1").type).toBe("empty")
		expect(again.sheet("S").style("B1")).toEqual(style)
	})

	it("flattens row/column DEFAULT styles into per-cell styles (documented)", async () => {
		// col-row-styles.xlsx styles bare cells via <col style> and <row s customFormat> defaults.
		// The bridge writes each cell's EFFECTIVE style directly; the rewritten file has no
		// defaults but every cell reads back with the same format as before.
		const before = await openXlsx(await loadFixture("col-row-styles.xlsx"))
		const after = await openXlsx(await rewrite(before))
		const sheet = after.sheet("Sheet1")
		expect(sheet.numberFormat("B1")).toBe("mm-dd-yy") // was column-default
		expect(sheet.cell("B1").type).toBe("date")
		expect(sheet.numberFormat("A3")).toBe("0.00%") // was row-default
		expect(sheet.style("A1")).toBeUndefined() // unstyled stays unstyled
	})

	it("rewrites an UNSTYLED workbook to byte-identical archives (with and without dates)", async () => {
		// The implicit date format round-trips through style() as {numberFormat:'mm-dd-yy'}, which
		// reverse-maps to the same built-in id 14 — so even date-bearing bare input reproduces the
		// exact bytes, not merely equivalent ones.
		for (const input of [
			{ sheets: [{ name: "S", rows: [["a", 1, true], [3.14]] }] },
			{ sheets: [{ name: "S", rows: [["x", new Date(Date.UTC(2020, 0, 1))]] }] },
		]) {
			const first = await writeXlsx(input)
			const second = await rewrite(await openXlsx(first))
			expect(Array.from(second)).toEqual(Array.from(first))
		}
	})
})

describe("bridge — corpus property", () => {
	it("every readable fixture round-trips losslessly OR fails typed — never bare, never silent", async () => {
		const dataDir = fileURLToPath(new URL("../../../../fixtures/data/", import.meta.url))
		const files = (await readdir(dataDir)).filter((f) => f.endsWith(".xlsx"))
		expect(files.length).toBeGreaterThan(5)
		let lossless = 0
		const typedFailures: string[] = []
		for (const file of files) {
			let before: Workbook
			try {
				before = await openXlsx(await loadFixture(file))
			} catch {
				continue // the intentionally-broken fixtures
			}
			let bytes: Uint8Array
			try {
				bytes = await rewrite(before)
			} catch (e) {
				// A tolerated read the writer can't represent must surface as a TYPED error.
				expect(e, file).toBeInstanceOf(XlsxError)
				expect((e as XlsxError).code, file).toBe("invalid-input")
				typedFailures.push(file)
				continue
			}
			const after = await openXlsx(bytes)
			expect(await styleSnapshot(after), file).toEqual(await styleSnapshot(before))
			lossless++
		}
		expect(lossless).toBeGreaterThan(5)
		// The only fixture allowed to refuse: the fuzz file whose cell ref is a 300-letter column
		// (kept faithfully by the tolerant reader, but addressable nowhere on a writable grid).
		expect(typedFailures).toEqual(["edge-overflow-col.xlsx"])
	})
})

describe("bridge — hostile files (review regressions)", () => {
	// Hand-craft a minimal workbook whose sheet1.xml is `sheetXml`, through the writer's own zip
	// layer — for inputs the tolerant reader accepts but no writer should ever have produced.
	async function craftWorkbook(sheetXml: string): Promise<Uint8Array> {
		const { writeZip } = await import("../zip")
		const enc = new TextEncoder()
		const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		return writeZip([
			{
				name: "[Content_Types].xml",
				data: enc.encode(
					`${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
				),
			},
			{
				name: "_rels/.rels",
				data: enc.encode(
					`${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/workbook.xml",
				data: enc.encode(
					`${decl}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
				),
			},
			{
				name: "xl/_rels/workbook.xml.rels",
				data: enc.encode(
					`${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/worksheets/sheet1.xml",
				data: enc.encode(
					`${decl}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetXml}</sheetData></worksheet>`,
				),
			},
		])
	}

	async function bridgeError(sheetXml: string): Promise<XlsxError> {
		const wb = await openXlsx(await craftWorkbook(sheetXml))
		const err = await workbookToInput(wb).then(
			() => null,
			(e) => e,
		)
		expect(err).toBeInstanceOf(XlsxError)
		expect((err as XlsxError).code).toBe("invalid-input")
		return err as XlsxError
	}

	it("refuses a cell beyond Excel's grid quickly and typed, instead of hanging", async () => {
		// A ref like A99999999999999 PARSES fine — the danger is that its row number becomes the
		// length of the rows array, which the writer then iterates.
		const started = Date.now()
		const err = await bridgeError(
			'<row r="99999999999999"><c r="A99999999999999"><v>1</v></c></row>',
		)
		expect(Date.now() - started).toBeLessThan(2000) // typed refusal, not an hours-long loop
		expect(err.message).toMatch(/grid position/)
	})

	it("refuses case-variant duplicate refs instead of silently dropping one value", async () => {
		// "A1" and "a1" are DISTINCT cells to the reader (cell() keys by the verbatim ref) but one
		// grid slot to a writer — last-wins placement would make the value 2 vanish with no error.
		const err = await bridgeError(
			'<row r="1"><c r="A1"><v>2</v></c><c r="a1"><v>1</v></c></row>',
		)
		expect(err.message).toMatch(/"A1" and "a1".*one grid position/)
	})

	it("keeps last-wins for SAME-spelling duplicate refs, matching the reader's cell()", async () => {
		const wb = await openXlsx(
			await craftWorkbook('<row r="1"><c r="A1"><v>2</v></c><c r="A1"><v>1</v></c></row>'),
		)
		expect(wb.sheet("S").cell("A1").value).toBe(1) // reader's own answer is last-wins
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(wb)))
		expect(rewritten.sheet("S").cell("A1").value).toBe(1) // bridge agrees
	})
})
