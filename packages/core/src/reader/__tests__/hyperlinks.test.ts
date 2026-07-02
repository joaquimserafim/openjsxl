import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { parseRels, type Relationship } from "../../ooxml"
import { openXlsx } from "../workbook"
import { parseHyperlinks } from "../worksheet"

// Hyperlinks (F2.3). `<hyperlink>` elements sit in a `<hyperlinks>` block; an `r:id` resolves
// to the external target through the worksheet's own rels, while `location`, `tooltip`, and
// `display` come straight off the element. Verified against real Excel/openpyxl output
// (hyperlinks.xlsx — calamine, MIT) which mixes all of these, including links that carry both
// an external target and an in-workbook location.

function relsOf(...entries: Array<[string, string]>): Map<string, Relationship> {
	const xml = entries
		.map(
			([id, target]) => `<Relationship Id="${id}" Target="${target}" TargetMode="External"/>`,
		)
		.join("")
	return parseRels(`<Relationships>${xml}</Relationships>`)
}

describe("Worksheet.hyperlinks — real fixture", () => {
	it("reads every link on the sheet with targets resolved through rels", async () => {
		const wb = await openXlsx(await loadFixture("hyperlinks.xlsx"))
		expect(wb.sheet("Links").hyperlinks).toEqual([
			{ ref: "A1", target: "https://github.com/tafia/calamine" },
			{ ref: "B1:C2", target: "mailto:foo@example.com", tooltip: "Email Foo" },
			{
				ref: "A2",
				target: "https://www.rust-lang.org/",
				tooltip: "Rust homepage",
				display: "Rust Programming Language",
			},
			{ ref: "A3", location: "'Sheet2'!B5", display: "Sheet2 B5" },
			{ ref: "A4", target: "file:///Book2.xlsx" },
			{ ref: "A5", target: "file:///..\\Sales\\Book2.xlsx" },
			{ ref: "A6", target: "file:///C:\\Temp\\Book1.xlsx" },
			{ ref: "A7", target: "file:///Book2.xlsx", location: "Sheet1!A1" },
			{ ref: "A8", target: "file:///Book2.xlsx", location: "'Sales Data'!A1:G5" },
		])
	})

	it("is empty for a sheet without links (and without a rels part)", async () => {
		const wb = await openXlsx(await loadFixture("hyperlinks.xlsx"))
		expect(wb.sheet("Sheet2").hyperlinks).toEqual([])
	})

	it("parses once and caches the result", async () => {
		const sheet = (await openXlsx(await loadFixture("hyperlinks.xlsx"))).sheet("Links")
		expect(sheet.hyperlinks).toBe(sheet.hyperlinks)
	})
})

describe("parseHyperlinks — units", () => {
	it("resolves an external target from the worksheet rels", () => {
		const xml = '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>'
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([
			{ ref: "A1", target: "https://example.com/" },
		])
	})

	it("keeps an in-workbook location with no target", () => {
		const xml = '<hyperlink ref="A3" location="Sheet2!B5" display="go"/>'
		expect(parseHyperlinks(xml)).toEqual([{ ref: "A3", location: "Sheet2!B5", display: "go" }])
	})

	it("leaves target absent when the r:id has no matching rel", () => {
		const xml = '<hyperlink ref="A1" r:id="rIdX" tooltip="t"/>'
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([
			{ ref: "A1", tooltip: "t" },
		])
	})

	it("leaves target absent when no rels are supplied", () => {
		expect(parseHyperlinks('<hyperlink ref="A1" r:id="rId1"/>')).toEqual([{ ref: "A1" }])
	})

	it("tolerates a namespace prefix on the element and a non-r:id relationship attribute", () => {
		const xml = '<x:hyperlink ref="A1" rel:id="rId1"/>'
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([
			{ ref: "A1", target: "https://example.com/" },
		])
	})

	it("skips a hyperlink with a missing or empty ref", () => {
		const xml = '<hyperlinks><hyperlink/><hyperlink ref=""/><hyperlink ref="A1"/></hyperlinks>'
		expect(parseHyperlinks(xml)).toEqual([{ ref: "A1" }])
	})

	it("treats an empty external target like an empty location: no destination (F4.6 review)", () => {
		// A crafted rels part can carry Target="". Surfacing it as target:"" would break the
		// round-trip contract: the writer normalizes empty destinations away, so re-reading the
		// rewritten file would silently lose the key instead of matching or failing typed.
		const xml = '<hyperlink ref="A1" r:id="rId1" location="S!B2"/>'
		expect(parseHyperlinks(xml, relsOf(["rId1", ""]))).toEqual([
			{ ref: "A1", location: "S!B2" },
		])
	})
})
