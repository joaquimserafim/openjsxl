import { createHash } from "node:crypto";
import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";

// F6.2 — Worksheet.images(). Two fixtures: a REAL openpyxl+Pillow file (validates parseDrawing on
// genuine drawingML — unprefixed anchors, package-absolute media targets, real PNG/JPEG bytes), and
// a hand-crafted edge file (shared media + every degrade path openpyxl won't produce).

const sha = (b: Uint8Array): string => createHash("sha256").update(b).digest("hex").slice(0, 16);
const text = (b: Uint8Array): string => new TextDecoder().decode(b);

describe("Worksheet.images — real openpyxl fixture", () => {
	it("reads both pictures with exact bytes, mime, 1-based anchors, and document order", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-images.xlsx"));
		const images = await wb.sheet("Pics").images();
		expect(images.length).toBe(2);

		// Document order: the JPEG twoCellAnchor precedes the PNG oneCellAnchor in the drawing.
		const [jpeg, png] = images;

		// twoCellAnchor JPEG spanning D4:F8 — OOXML 0-based (3,3)→(5,7) reads 1-based (4,4)→(6,8).
		expect(jpeg?.mime).toBe("image/jpeg");
		expect(jpeg?.name).toBe("Image 2");
		expect(jpeg?.anchor).toEqual({
			from: { col: 4, row: 4, colOff: 0, rowOff: 0 },
			to: { col: 6, row: 8, colOff: 0, rowOff: 0 },
			editAs: "oneCell",
		});
		expect(sha(jpeg?.bytes as Uint8Array)).toBe("0a78a635f52a1b65"); // matches the source JPEG

		// oneCellAnchor PNG pinned at B2 with EMU offsets + a fixed extent.
		expect(png?.mime).toBe("image/png");
		expect(png?.name).toBe("Image 1");
		expect(png?.anchor).toEqual({
			from: { col: 2, row: 2, colOff: 9525, rowOff: 19050 },
			ext: { cx: 762000, cy: 571500 },
		});
		expect(sha(png?.bytes as Uint8Array)).toBe("cacab3c9385b24aa"); // matches the source PNG
	});

	it("caches: repeated calls return the same array (one media resolution)", async () => {
		const sheet = (await openXlsx(await loadFixture("openpyxl-images.xlsx"))).sheet("Pics");
		const first = await sheet.images();
		const second = await sheet.images();
		expect(second).toBe(first);
	});
});

describe("Worksheet.images — crafted edge fixture", () => {
	it("shares one bytes buffer between pictures on the same media part", async () => {
		const wb = await openXlsx(await loadFixture("images-edge.xlsx"));
		const images = await wb.sheet("Pics").images();
		// Only A and B survive; both reference the same media part.
		expect(images.map((i) => i.name)).toEqual(["A", "B"]);
		expect(images[0]?.bytes).toBe(images[1]?.bytes); // same buffer, read once
		expect(text(images[0]?.bytes as Uint8Array)).toBe("openjsxl-shared-media-bytes");
		expect(images.every((i) => i.mime === "image/png")).toBe(true);
		expect(images[0]?.anchor).toEqual({
			from: { col: 1, row: 1, colOff: 0, rowOff: 0 },
			ext: { cx: 100, cy: 100 },
		});
		expect(images[1]?.anchor).toEqual({
			from: { col: 3, row: 3, colOff: 0, rowOff: 0 },
			to: { col: 5, row: 5, colOff: 0, rowOff: 0 },
		});
	});

	it("degrades every unreadable picture rather than throwing", async () => {
		// The fixture also holds an absoluteAnchor picture (C), a shape (D), a picture whose media
		// part is absent (E), and one whose embed isn't in the drawing rels (F) — all dropped.
		const wb = await openXlsx(await loadFixture("images-edge.xlsx"));
		const images = await wb.sheet("Pics").images();
		expect(images.length).toBe(2);
		// The rest of the sheet still reads normally.
		expect(wb.sheet("Pics").cell("A1").value).toBe("pics");
	});
});

describe("Worksheet.images — sheets without pictures", () => {
	it("returns [] for a sheet with no drawing (and doesn't disturb other accessors)", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"));
		expect(await wb.sheet("Sheet1").images()).toEqual([]);
		expect(wb.sheet("Sheet1").cell("A1").value).toBe("hello");
	});
});

// M6-analysis coverage: structural rel-graph edges no real producer emits — crafted through the
// writer's own zip layer, like the bridge's hostile-file suite.
describe("Worksheet.images — crafted rel-graph edges", () => {
	const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
	const NS_XDR =
		'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
	const REL_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
	const TYPE_DRAWING =
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
	const TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

	const drawingXml = (name: string): string =>
		`${DECL}<xdr:wsDr ${NS_XDR}><xdr:oneCellAnchor>` +
		`<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
		`<xdr:ext cx="100" cy="100"/>` +
		`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="${name}"/></xdr:nvPicPr>` +
		`<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic>` +
		`<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;

	// A minimal workbook whose sheet has the given drawing rels + drawing parts.
	async function craft(
		sheetDrawingRels: string,
		drawings: readonly { name: string; xml: string; rels: string }[],
	): Promise<Uint8Array> {
		const { writeZip } = await import("../../writer/zip");
		const enc = new TextEncoder();
		return writeZip([
			{
				name: "[Content_Types].xml",
				data: enc.encode(
					`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
				),
			},
			{
				name: "_rels/.rels",
				data: enc.encode(
					`${DECL}<Relationships ${REL_NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/workbook.xml",
				data: enc.encode(
					`${DECL}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
				),
			},
			{
				name: "xl/_rels/workbook.xml.rels",
				data: enc.encode(
					`${DECL}<Relationships ${REL_NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/worksheets/sheet1.xml",
				data: enc.encode(
					`${DECL}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
				),
			},
			{
				name: "xl/worksheets/_rels/sheet1.xml.rels",
				data: enc.encode(
					`${DECL}<Relationships ${REL_NS}>${sheetDrawingRels}</Relationships>`,
				),
			},
			...drawings.flatMap((d) => [
				{ name: `xl/drawings/${d.name}.xml`, data: enc.encode(d.xml) },
				{
					name: `xl/drawings/_rels/${d.name}.xml.rels`,
					data: enc.encode(`${DECL}<Relationships ${REL_NS}>${d.rels}</Relationships>`),
				},
			]),
			{ name: "xl/media/image1.png", data: enc.encode("crafted-media-bytes") },
		]);
	}

	it("reads pictures from MULTIPLE /drawing rels on one sheet, in rel order", async () => {
		// Valid OOXML a real producer rarely emits: sheet1 → drawing1.xml AND drawing2.xml, each
		// holding one picture that shares the same media part.
		const mediaRel = `<Relationship Id="rId1" Type="${TYPE_IMAGE}" Target="../media/image1.png"/>`;
		const bytes = await craft(
			`<Relationship Id="rId1" Type="${TYPE_DRAWING}" Target="../drawings/drawing1.xml"/>` +
				`<Relationship Id="rId2" Type="${TYPE_DRAWING}" Target="../drawings/drawing2.xml"/>`,
			[
				{ name: "drawing1", xml: drawingXml("first"), rels: mediaRel },
				{ name: "drawing2", xml: drawingXml("second"), rels: mediaRel },
			],
		);
		const images = await (await openXlsx(bytes)).sheet("S").images();
		expect(images.map((i) => i.name)).toEqual(["first", "second"]); // neither drawing dropped
		expect(images.every((i) => text(i.bytes) === "crafted-media-bytes")).toBe(true);
	});

	it("skips a picture whose media rel is TargetMode=External (no throw, others kept)", async () => {
		const bytes = await craft(
			`<Relationship Id="rId1" Type="${TYPE_DRAWING}" Target="../drawings/drawing1.xml"/>`,
			[
				{
					name: "drawing1",
					xml: drawingXml("linked"),
					rels: `<Relationship Id="rId1" Type="${TYPE_IMAGE}" Target="https://example.com/logo.png" TargetMode="External"/>`,
				},
			],
		);
		const wb = await openXlsx(bytes);
		expect(await wb.sheet("S").images()).toEqual([]); // linked-not-embedded → degrade, not throw
	});
});
