import { describe, expect, it } from "vitest"
import { parseStyles } from "../styles"

// F4.1 — the full style read model: fonts, fills, borders, alignment, colors, and the cached
// cellStyle(i) materializer. Inline styleSheet XML keeps each case surgical; end-to-end coverage
// against a real producer lives in reader/__tests__/cell-style.test.ts.

const SHEET = (inner: string): string =>
	`<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${inner}</styleSheet>`

// A minimal three-table prelude: font 0 / fill 0 (none) / border 0 (empty) are the workbook
// defaults every real file carries; xf 0 is the default cell format.
const DEFAULTS = {
	fonts: '<font><sz val="11"/><name val="Calibri"/></font>',
	fills: '<fill><patternFill patternType="none"/></fill>',
	borders: "<border><left/><right/><top/><bottom/><diagonal/></border>",
	xf0: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
}

describe("parseStyles — cellStyle materialization", () => {
	it("resolves the default xf (and out-of-range indexes) to undefined", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="1">${DEFAULTS.fonts}</fonts>` +
					`<fills count="1">${DEFAULTS.fills}</fills>` +
					`<borders count="1">${DEFAULTS.borders}</borders>` +
					`<cellXfs count="1">${DEFAULTS.xf0}</cellXfs>`,
			),
		)
		expect(table.cellStyle(0)).toBeUndefined()
		expect(table.cellStyle(undefined)).toBeUndefined() // omitted s = style 0
		expect(table.cellStyle(99)).toBeUndefined()
	})

	it("reads fonts: name, size, bold, italic, underline, strike, rgb color", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="2">${DEFAULTS.fonts}` +
					'<font><name val="Arial"/><sz val="14"/><b/><i val="1"/><u/><strike/><color rgb="FFFF0000"/></font>' +
					"</fonts>" +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)).toEqual({
			font: {
				name: "Arial",
				size: 14,
				bold: true,
				italic: true,
				underline: "single",
				strike: true,
				color: { rgb: "FFFF0000" },
			},
		})
	})

	it("keeps theme+tint, indexed, and auto colors raw", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="4">${DEFAULTS.fonts}` +
					'<font><color theme="4" tint="0.4"/></font>' +
					'<font><color indexed="10"/></font>' +
					'<font><color auto="1"/></font>' +
					"</fonts>" +
					`<cellXfs count="4">${DEFAULTS.xf0}` +
					'<xf fontId="1"/><xf fontId="2"/><xf fontId="3"/></cellXfs>',
			),
		)
		expect(table.cellStyle(1)?.font?.color).toEqual({ theme: 4, tint: 0.4 })
		expect(table.cellStyle(2)?.font?.color).toEqual({ indexed: 10 })
		expect(table.cellStyle(3)?.font?.color).toEqual({ auto: true })
	})

	it("reads pattern fills and treats the solid fgColor as the visible color", () => {
		const table = parseStyles(
			SHEET(
				`<fills count="3">${DEFAULTS.fills}` +
					'<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>' +
					'<fill><patternFill patternType="lightGray"><fgColor theme="2"/></patternFill></fill>' +
					"</fills>" +
					`<cellXfs count="3">${DEFAULTS.xf0}<xf fillId="1"/><xf fillId="2"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)?.fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FFFFFF00" },
			bgColor: { indexed: 64 },
		})
		expect(table.cellStyle(2)?.fill).toEqual({
			patternType: "lightGray",
			fgColor: { theme: 2 },
		})
	})

	it("a gradient fill degrades to no fill (documented deferral)", () => {
		const table = parseStyles(
			SHEET(
				`<fills count="2">${DEFAULTS.fills}` +
					'<fill><gradientFill degree="90"><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill>' +
					"</fills>" +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf fillId="1"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)).toBeUndefined()
	})

	it("reads per-edge borders, with and without colors, ignoring styleless edges", () => {
		const table = parseStyles(
			SHEET(
				`<borders count="2">${DEFAULTS.borders}` +
					'<border><left style="dashed"><color rgb="FF0070C0"/></left><right/>' +
					'<top style="thin"/><bottom style="double"/></border>' +
					"</borders>" +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf borderId="1"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)?.border).toEqual({
			left: { style: "dashed", color: { rgb: "FF0070C0" } },
			top: { style: "thin" },
			bottom: { style: "double" },
		})
	})

	it("reads inline alignment and drops the legacy textRotation=255 marker", () => {
		const table = parseStyles(
			SHEET(
				`<cellXfs count="3">${DEFAULTS.xf0}` +
					'<xf numFmtId="0"><alignment horizontal="center" vertical="top" wrapText="1" shrinkToFit="1" indent="2" textRotation="45"/></xf>' +
					'<xf numFmtId="0"><alignment textRotation="255"/></xf>' +
					"</cellXfs>",
			),
		)
		expect(table.cellStyle(1)?.alignment).toEqual({
			horizontal: "center",
			vertical: "top",
			wrapText: true,
			shrinkToFit: true,
			indent: 2,
			textRotation: 45,
		})
		// 255 is "vertical stacked", not degrees — with nothing else set, the xf has no style.
		expect(table.cellStyle(2)).toBeUndefined()
	})

	it("includes the number format code for non-General ids", () => {
		const table = parseStyles(
			SHEET(
				'<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>' +
					`<cellXfs count="3">${DEFAULTS.xf0}<xf numFmtId="10"/><xf numFmtId="164"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)).toEqual({ numberFormat: "0.00%" })
		expect(table.cellStyle(2)).toEqual({ numberFormat: "yyyy-mm-dd" })
	})

	it("is reference-stable: the same index returns the same object", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="2">${DEFAULTS.fonts}<font><b/></font></fonts>` +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf fontId="1"/></cellXfs>`,
			),
		)
		expect(table.cellStyle(1)).toBe(table.cellStyle(1))
	})

	it("degrades garbage gracefully: unknown enums, bad numerics, accounting underline", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="3">${DEFAULTS.fonts}` +
					'<font><u val="singleAccounting"/><sz val="banana"/></font>' +
					'<font><color theme="x"/></font>' +
					"</fonts>" +
					`<fills count="2">${DEFAULTS.fills}<fill><patternFill patternType="plaid"/></fill></fills>` +
					`<borders count="2">${DEFAULTS.borders}<border><top style="wavy"/></border></borders>` +
					`<cellXfs count="4">${DEFAULTS.xf0}` +
					'<xf fontId="1"/><xf fontId="2" fillId="1" borderId="1"/>' +
					'<xf><alignment horizontal="middle"/></xf>' +
					"</cellXfs>",
			),
		)
		// Accounting underline + unparseable size leave an empty font → no style at all.
		expect(table.cellStyle(1)).toBeUndefined()
		// Bad theme number → no color; unknown pattern → 'none' → no fill; unknown border style
		// → no edge → no border. Everything collapses to undefined.
		expect(table.cellStyle(2)).toBeUndefined()
		// Unknown horizontal value → no alignment.
		expect(table.cellStyle(3)).toBeUndefined()
	})

	// Regressions — misnested sections must not poison later parsing (adversarial review, F4.1).
	// The tokenizer is non-validating, so an unclosed record's builder used to leak past its
	// section and either swallow structural tokens or capture <dxf> look-alike children.

	it("an unclosed <font> neither disables the hot path nor swallows cellXfs", () => {
		const table = parseStyles(
			SHEET(
				'<fonts count="1"><font><sz val="11"/></fonts>' + // <font> never closed
					'<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>',
			),
		)
		// Pre-fix: the dangling font builder intercepted <cellXfs>/<xf>, so xfs stayed empty and
		// even date detection (the pre-F4.1 behavior) regressed. Both must survive.
		expect(table.isDateStyle(1)).toBe(true)
		expect(table.formatCode(1)).toBe("mm-dd-yy")
		expect(table.cellStyle(1)).toEqual({ numberFormat: "mm-dd-yy" })
	})

	it("an unclosed <fill> is flushed at </fills>; a later dxf fill is not captured", () => {
		const table = parseStyles(
			SHEET(
				`<fills count="2">${DEFAULTS.fills}` +
					'<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill>' + // unclosed
					"</fills>" +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf fillId="1"/></cellXfs>` +
					'<dxfs count="1"><dxf><fill><patternFill patternType="darkTrellis"><fgColor rgb="FFBADBAD"/></patternFill></fill></dxf></dxfs>',
			),
		)
		// The sheet's own solid green lands at index 1; the conditional-format fill must not.
		expect(table.cellStyle(1)?.fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FF00FF00" },
		})
	})

	it("an unclosed <border> edge is flushed at </borders>; dxf colors are not grafted on", () => {
		const table = parseStyles(
			SHEET(
				`<borders count="2">${DEFAULTS.borders}` +
					'<border><left style="thin">' + // border and its edge both unclosed
					"</borders>" +
					`<cellXfs count="2">${DEFAULTS.xf0}<xf borderId="1"/></cellXfs>` +
					'<dxfs count="1"><dxf><font><color rgb="FFDD0000"/></font><border></border></dxf></dxfs>',
			),
		)
		// The dangling edge commits as authored (thin, no color) — the dxf font color must not
		// become its color, and the dxf border must not append a phantom record.
		expect(table.cellStyle(1)?.border).toEqual({ left: { style: "thin" } })
	})

	it("well-formed dxf blocks never contaminate the component tables", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="1">${DEFAULTS.fonts}</fonts>` +
					`<fills count="1">${DEFAULTS.fills}</fills>` +
					`<borders count="1">${DEFAULTS.borders}</borders>` +
					`<cellXfs count="1">${DEFAULTS.xf0}</cellXfs>` +
					'<dxfs count="1"><dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFBADBAD"/></patternFill></fill></dxf></dxfs>',
			),
		)
		// dxf children live outside the fonts/fills sections and must not extend the tables.
		expect(table.cellStyle(0)).toBeUndefined()
		expect(table.cellStyle(1)).toBeUndefined()
	})

	it("ignores cellStyleXfs — a cell s indexes cellXfs only", () => {
		const table = parseStyles(
			SHEET(
				`<fonts count="2">${DEFAULTS.fonts}<font><b/></font></fonts>` +
					'<cellStyleXfs count="2"><xf fontId="1"/><xf fontId="1"/></cellStyleXfs>' +
					`<cellXfs count="1">${DEFAULTS.xf0}</cellXfs>`,
			),
		)
		expect(table.cellStyle(0)).toBeUndefined()
		expect(table.cellStyle(1)).toBeUndefined() // cellStyleXfs entries must not leak in
	})
})
