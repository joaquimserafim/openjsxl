import { describe, expect, it } from "vitest";
import { type DrawingImage, mimeForMediaPath, parseDrawing } from "../drawing";

// F6.2 — the spreadsheetDrawing parser. Excel/openpyxl put the anchor elements in the DEFAULT
// (unprefixed) namespace and use a:/r: only on the blip, so these samples mirror that shape. The
// parser is pure: it returns the unresolved r:embed id; the reader does the media I/O.

const NS =
	'xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const pic = (id: number, name: string, embed: string) =>
	`<pic><nvPicPr><cNvPr id="${id}" name="${name}"/><cNvPicPr/></nvPicPr><blipFill><a:blip r:embed="${embed}"/></blipFill><spPr/></pic><clientData/>`;
const wsDr = (inner: string) => `<?xml version="1.0"?>\n<wsDr ${NS}>${inner}</wsDr>`;

describe("parseDrawing", () => {
	it("reads a oneCellAnchor picture: 1-based cell, EMU offsets + ext, name, embed", () => {
		const xml = wsDr(
			`<oneCellAnchor><from><col>1</col><colOff>9525</colOff><row>2</row><rowOff>19050</rowOff></from>` +
				`<ext cx="762000" cy="571500"/>${pic(1, "Logo", "rId7")}</oneCellAnchor>`,
		);
		expect(parseDrawing(xml)).toEqual<DrawingImage[]>([
			{
				embed: "rId7",
				name: "Logo",
				from: { col: 2, row: 3, colOff: 9525, rowOff: 19050 }, // 0-based 1/2 → 1-based 2/3
				ext: { cx: 762000, cy: 571500 },
			},
		]);
	});

	it("reads a twoCellAnchor picture with from/to and editAs", () => {
		const xml = wsDr(
			`<twoCellAnchor editAs="oneCell"><from><col>3</col><colOff>0</colOff><row>3</row><rowOff>0</rowOff></from>` +
				`<to><col>5</col><colOff>0</colOff><row>7</row><rowOff>0</rowOff></to>${pic(2, "X", "rId2")}</twoCellAnchor>`,
		);
		expect(parseDrawing(xml)).toEqual<DrawingImage[]>([
			{
				embed: "rId2",
				name: "X",
				from: { col: 4, row: 4, colOff: 0, rowOff: 0 },
				to: { col: 6, row: 8, colOff: 0, rowOff: 0 },
				editAs: "oneCell",
			},
		]);
	});

	it("skips absoluteAnchor pictures and non-picture (shape) anchors", () => {
		const xml = wsDr(
			`<absoluteAnchor><pos x="0" y="0"/><ext cx="1" cy="1"/>${pic(1, "abs", "rId1")}</absoluteAnchor>` +
				`<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>` +
				`<to><col>1</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></to>` +
				`<sp><nvSpPr><cNvPr id="2" name="shape"/></nvSpPr><spPr/></sp><clientData/></twoCellAnchor>`,
		);
		expect(parseDrawing(xml)).toEqual([]);
	});

	it("skips a picture with no resolvable blip embed", () => {
		const xml = wsDr(
			`<oneCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>` +
				`<ext cx="1" cy="1"/><pic><nvPicPr><cNvPr id="1" name="noblip"/></nvPicPr><blipFill/><spPr/></pic><clientData/></oneCellAnchor>`,
		);
		expect(parseDrawing(xml)).toEqual([]);
	});

	it("keeps document order and returns one entry per picture anchor", () => {
		const xml = wsDr(
			`<oneCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><ext cx="1" cy="1"/>${pic(1, "first", "rIdA")}</oneCellAnchor>` +
				`<oneCellAnchor><from><col>0</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx="1" cy="1"/>${pic(2, "second", "rIdB")}</oneCellAnchor>`,
		);
		expect(parseDrawing(xml).map((d) => [d.name, d.embed])).toEqual([
			["first", "rIdA"],
			["second", "rIdB"],
		]);
	});

	it("tolerates junk integers (non-numeric → 0) and never throws", () => {
		const xml = wsDr(
			`<oneCellAnchor><from><col>oops</col><colOff></colOff><row/><rowOff>x</rowOff></from>` +
				`<ext cx="nope" cy="12"/>${pic(1, "junk", "rId1")}</oneCellAnchor>`,
		);
		expect(parseDrawing(xml)).toEqual<DrawingImage[]>([
			{
				embed: "rId1",
				name: "junk",
				from: { col: 1, row: 1, colOff: 0, rowOff: 0 },
				ext: { cx: 0, cy: 12 },
			},
		]);
	});

	it("returns [] for an empty or anchor-less drawing", () => {
		expect(parseDrawing(wsDr(""))).toEqual([]);
		expect(parseDrawing("not even xml")).toEqual([]);
	});

	// Review regression: a picture's spPr/a:xfrm/a:ext (the shape-transform extent Excel and
	// LibreOffice emit) must NOT overwrite the anchor's own <ext>, nor add an ext to a two-cell anchor.
	it("ignores the picture's spPr/xfrm ext (keeps the anchor's own extent)", () => {
		const one = wsDr(
			`<oneCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>` +
				`<ext cx="990000" cy="792000"/>` +
				`<pic><nvPicPr><cNvPr id="1" name="Pic"/></nvPicPr><blipFill><a:blip r:embed="rId1"/></blipFill>` +
				`<spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12345" cy="67890"/></a:xfrm></spPr></pic></oneCellAnchor>`,
		);
		expect(parseDrawing(one)[0]?.ext).toEqual({ cx: 990000, cy: 792000 }); // anchor ext, not shape

		const two = wsDr(
			`<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>` +
				`<to><col>3</col><colOff>0</colOff><row>3</row><rowOff>0</rowOff></to>` +
				`<pic><nvPicPr><cNvPr id="1" name="Pic"/></nvPicPr><blipFill><a:blip r:embed="rId1"/></blipFill>` +
				`<spPr><a:xfrm><a:ext cx="999" cy="888"/></a:xfrm></spPr></pic></twoCellAnchor>`,
		);
		expect(parseDrawing(two)[0]?.ext).toBeUndefined(); // a two-cell anchor must carry no ext
	});

	// Milestone review (F6.4) regression: the anchor ELEMENT names its shape — a oneCellAnchor is
	// positioned by <ext>, a twoCellAnchor by <to>. The parser must return EXACTLY ONE of the two,
	// matching what the writer accepts, so a malformed anchor can't produce a record the writer then
	// rejects (which would nuke a whole read→write rewrite). Keep the field the kind calls for.
	it("normalizes a malformed anchor to exactly one of to/ext (kind is authoritative)", () => {
		// A oneCellAnchor whose required <ext> is missing can't be placed — drop it (degrade).
		const noExt = wsDr(
			`<oneCellAnchor><from><col>0</col><row>0</row></from>${pic(1, "noext", "rId1")}</oneCellAnchor>`,
		);
		expect(parseDrawing(noExt)).toEqual([]);

		// A twoCellAnchor carrying a stray <ext> before its <pic> keeps its <to> and drops the ext —
		// it's a valid two-cell picture, so recover it rather than reject the whole file.
		const strayExt = wsDr(
			`<twoCellAnchor><from><col>0</col><row>0</row></from><to><col>2</col><row>2</row></to>` +
				`<ext cx="55" cy="55"/>${pic(1, "two", "rId1")}</twoCellAnchor>`,
		);
		const two = parseDrawing(strayExt);
		expect(two[0]?.to).toEqual({ col: 3, row: 3, colOff: 0, rowOff: 0 });
		expect(two[0]?.ext).toBeUndefined();

		// A oneCellAnchor carrying BOTH a <to> and an <ext> keeps the ext (its kind) and drops the to.
		const both = wsDr(
			`<oneCellAnchor><from><col>0</col><row>0</row></from><to><col>2</col><row>2</row></to>` +
				`<ext cx="77" cy="77"/>${pic(1, "one", "rId1")}</oneCellAnchor>`,
		);
		const one = parseDrawing(both);
		expect(one[0]?.ext).toEqual({ cx: 77, cy: 77 });
		expect(one[0]?.to).toBeUndefined();
	});

	// Review regression: a picture nested in a <grpSp> carries the group's anchor, not its own —
	// skip it (and never let the group's cNvPr name leak onto a picture).
	it("skips a picture nested inside a group shape (grpSp)", () => {
		const grouped = wsDr(
			`<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>` +
				`<to><col>2</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></to>` +
				`<grpSp><nvGrpSpPr><cNvPr id="9" name="Group 9"/></nvGrpSpPr>` +
				`<pic><nvPicPr><cNvPr id="1" name="Inner"/></nvPicPr><blipFill><a:blip r:embed="rId1"/></blipFill><spPr/></pic>` +
				`</grpSp></twoCellAnchor>`,
		);
		expect(parseDrawing(grouped)).toEqual([]);
	});
});

describe("mimeForMediaPath", () => {
	it("maps known image extensions and degrades unknown ones", () => {
		expect(mimeForMediaPath("xl/media/image1.png")).toBe("image/png");
		expect(mimeForMediaPath("xl/media/image2.jpeg")).toBe("image/jpeg");
		expect(mimeForMediaPath("x.JPG")).toBe("image/jpeg"); // case-insensitive
		expect(mimeForMediaPath("x.gif")).toBe("image/gif");
		expect(mimeForMediaPath("x.emf")).toBe("image/x-emf");
		expect(mimeForMediaPath("x.weird")).toBe("application/octet-stream");
		expect(mimeForMediaPath("noext")).toBe("application/octet-stream");
	});
});
