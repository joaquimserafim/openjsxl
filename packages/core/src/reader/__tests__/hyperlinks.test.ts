import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { parseRels, parseWorkbook, type Relationship } from "../../ooxml";
import { writeXlsx } from "../../writer/workbook";
import { openXlsx } from "../workbook";
import { parseHyperlinks } from "../worksheet";

// Hyperlinks (F2.3). `<hyperlink>` elements sit in a `<hyperlinks>` block; an `r:id` resolves
// to the external target through the worksheet's own rels, while `location` is decoded once as
// ST_Xstring and `tooltip`/`display` come straight off the element. Verified against real
// Excel/openpyxl output (hyperlinks.xlsx — calamine, MIT), which mixes all of these, including
// links that carry both an external target and a location.

function relsOf(...entries: Array<[string, string]>): Map<string, Relationship> {
	const xml = entries
		.map(
			([id, target]) => `<Relationship Id="${id}" Target="${target}" TargetMode="External"/>`,
		)
		.join("");
	return parseRels(`<Relationships>${xml}</Relationships>`);
}

describe("Worksheet.hyperlinks — real fixture", () => {
	it("reads every link on the sheet with targets resolved through rels", async () => {
		const wb = await openXlsx(await loadFixture("hyperlinks.xlsx"));
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
		]);
	});

	it("is empty for a sheet without links (and without a rels part)", async () => {
		const wb = await openXlsx(await loadFixture("hyperlinks.xlsx"));
		expect(wb.sheet("Sheet2").hyperlinks).toEqual([]);
	});

	it("parses once and caches the result", async () => {
		const sheet = (await openXlsx(await loadFixture("hyperlinks.xlsx"))).sheet("Links");
		expect(sheet.hyperlinks).toBe(sheet.hyperlinks);
	});
});

describe("parseHyperlinks — units", () => {
	it("resolves an external target from the worksheet rels", () => {
		const xml = '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>';
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([
			{ ref: "A1", target: "https://example.com/" },
		]);
	});

	it("keeps an in-workbook location with no target", () => {
		const xml = '<hyperlink ref="A3" location="Sheet2!B5" display="go"/>';
		expect(parseHyperlinks(xml)).toEqual([{ ref: "A3", location: "Sheet2!B5", display: "go" }]);
	});

	it("drops a destination-less link when the r:id has no matching rel", () => {
		const xml = '<hyperlink ref="A1" r:id="rIdX" tooltip="t"/>';
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([]);
	});

	it("drops a destination-less link when no rels are supplied", () => {
		expect(parseHyperlinks('<hyperlink ref="A1" r:id="rId1"/>')).toEqual([]);
	});

	it("tolerates a namespace prefix on the element and a non-r:id relationship attribute", () => {
		const xml = '<x:hyperlink ref="A1" rel:id="rId1"/>';
		expect(parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]))).toEqual([
			{ ref: "A1", target: "https://example.com/" },
		]);
	});

	it("skips a hyperlink with a missing or empty ref", () => {
		const xml =
			'<hyperlinks><hyperlink location="S!A1"/><hyperlink ref="" location="S!A1"/>' +
			'<hyperlink ref="A1" location="S!A1"/></hyperlinks>';
		expect(parseHyperlinks(xml)).toEqual([{ ref: "A1", location: "S!A1" }]);
	});

	it("treats an empty external target like an empty location: no destination (F4.6 review)", () => {
		// A crafted rels part can carry Target="". Surfacing it as target:"" would break the
		// round-trip contract: the writer normalizes empty destinations away, so re-reading the
		// rewritten file would silently lose the key instead of matching or failing typed.
		const xml = '<hyperlink ref="A1" r:id="rId1" location="S!B2"/>';
		expect(parseHyperlinks(xml, relsOf(["rId1", ""]))).toEqual([
			{ ref: "A1", location: "S!B2" },
		]);
	});

	it("decodes each complete location exactly once after XML entity decoding", () => {
		const xml =
			"<hyperlinks>" +
			'<hyperlink ref="A1" location="_x0041_"/>' + // raw escape → A
			'<hyperlink ref="A2" location="_x005F_x0041_"/>' + // protected literal
			'<hyperlink ref="A3" location="_x005f_x0041_"/>' + // Excel Online lowercase quad
			'<hyperlink ref="A4" location="_x0041_x0042_"/>' + // shared underscore consumes once
			'<hyperlink ref="A5" location="&apos;R&amp;D _x005F_x0041_&apos;!A1"/>' +
			'<hyperlink ref="A6" location="row_x000D_two"/>' +
			'<hyperlink ref="A7" location="tab&#9;line&#10;end"/>' +
			'<hyperlink ref="A8" location="_x005F_x0001_"/>' + // protected control shape stays text
			'<hyperlink ref="A9" location="&#95;x0041_"/>' + // entity first, then ST_Xstring → A
			"</hyperlinks>";
		expect(parseHyperlinks(xml)).toEqual([
			{ ref: "A1", location: "A" },
			{ ref: "A2", location: "_x0041_" },
			{ ref: "A3", location: "_x0041_" },
			{ ref: "A4", location: "Ax0042_" },
			{ ref: "A5", location: "'R&D _x0041_'!A1" },
			{ ref: "A6", location: "row\rtwo" },
			{ ref: "A7", location: "tab\tline\nend" },
			{ ref: "A8", location: "_x0001_" },
			{ ref: "A9", location: "A" },
		]);
	});

	it("degrades unsafe decoded locations and keeps reader output inside writer bounds", async () => {
		const xml =
			"<hyperlinks>" +
			'<hyperlink ref="A1" r:id="rId1" location="_x0001_" tooltip="kept"/>' +
			'<hyperlink ref="A2" location="_x0001_" display="dropped with link"/>' +
			'<hyperlink ref="A3" r:id="rId1" location="_xD800_"/>' +
			'<hyperlink ref="A4" location="&#1;"/>' +
			'<hyperlink ref="A5" location="safe_x000D_cr"/>' +
			'<hyperlink ref="A6" location="_xFFFE_"/>' +
			'<hyperlink ref="A7" r:id="rId1" location="_x0041_"/>' +
			"</hyperlinks>";
		const hyperlinks = parseHyperlinks(xml, relsOf(["rId1", "https://example.com/"]));
		expect(hyperlinks).toEqual([
			{ ref: "A1", target: "https://example.com/", tooltip: "kept" },
			{ ref: "A3", target: "https://example.com/" },
			{ ref: "A5", location: "safe\rcr" },
			{ ref: "A7", target: "https://example.com/", location: "A" },
		]);

		// The degraded public model can cross the normal reader → writer boundary without the
		// strict writer rejecting a decoded control character or a destination-less record.
		const rewritten = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [["links"]], hyperlinks }] }),
		);
		expect(rewritten.sheet("S").hyperlinks).toEqual(hyperlinks);
	});

	it("decodes both raw tab and raw location spellings to the same logical sheet token", () => {
		// Tolerant producer-shaped direction: older files can carry both fields raw.
		const [sheet] = parseWorkbook(
			'<workbook><sheets><sheet name="_x0041_" r:id="rId1"/></sheets></workbook>',
		).sheets;
		const [link] = parseHyperlinks('<hyperlink ref="A1" location="&apos;_x0041_&apos;!A1"/>');
		expect(sheet?.name).toBe("A");
		expect(link?.location).toBe("'A'!A1");
	});
});
