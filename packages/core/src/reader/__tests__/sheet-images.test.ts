import { describe, expect, it } from "vitest";
import { parseRels } from "../../ooxml";
import { openZip } from "../../zip";
import { decodeText } from "../decode";
import { cachedRead, claimDrawings, loadSheetImages } from "../workbook";
import { DRAWING_REL, pictureBytes, pictureWorkbook } from "./picture-fixtures";
import { rewriteParts } from "./shared-part-fixtures";

// F11.96: the picture loader's parts. claimDrawings decides at open which sheet owns each drawing;
// cachedRead shares one read per image part across the workbook; loadSheetImages reads through both.

const SHEET = "xl/worksheets/sheet1.xml";
const SHEET_RELS = "xl/worksheets/_rels/sheet1.xml.rels";
const DRAWING = "xl/drawings/drawing1.xml";
const DRAWING_RELS = "xl/drawings/_rels/drawing1.xml.rels";
const DRAWING2 = "xl/drawings/drawing2.xml";
const MEDIA = "xl/media/image1.png";

// S1's rels name, in order: S2's drawing, an external drawing, S1's own drawing (the writer's rId1,
// moved last), its own drawing again, and a missing drawing.
const crafted = await rewriteParts(await pictureWorkbook([pictureBytes(1), pictureBytes(2)]), {
	[SHEET_RELS]: (t) =>
		t
			.replace(
				'<Relationship Id="rId1"',
				`<Relationship Id="rId7" Type="${DRAWING_REL}" Target="../drawings/drawing2.xml"/><Relationship Id="rId1"`,
			)
			.replace(
				"</Relationships>",
				`<Relationship Id="rId8" Type="${DRAWING_REL}" Target="https://example.com/d.xml" TargetMode="External"/>` +
					`<Relationship Id="rId9" Type="${DRAWING_REL}" Target="../drawings/drawing1.xml"/>` +
					`<Relationship Id="rId10" Type="${DRAWING_REL}" Target="../drawings/missing.xml"/></Relationships>`,
			),
});
const zip = openZip(crafted);
const rels = parseRels(decodeText(await zip.read(SHEET_RELS)));

describe("claimDrawings", () => {
	it("claims each existing internal drawing part and its rels once, in relationship order", () => {
		const claimed = new Set<string>();
		expect(claimDrawings(zip, SHEET, rels, claimed)).toEqual([
			{ path: DRAWING2, rels: "xl/drawings/_rels/drawing2.xml.rels" },
			{ path: DRAWING, rels: DRAWING_RELS },
		]);
		expect([...claimed]).toEqual([
			DRAWING2,
			"xl/drawings/_rels/drawing2.xml.rels",
			DRAWING,
			DRAWING_RELS,
		]);
	});

	it("owns nothing an earlier sheet claimed, and no rels part another role claimed", () => {
		expect(claimDrawings(zip, SHEET, rels, new Set([DRAWING, DRAWING2]))).toEqual([]);
		expect(claimDrawings(zip, SHEET, rels, new Set([DRAWING2, DRAWING_RELS]))).toEqual([
			{ path: DRAWING, rels: undefined },
		]);
		expect(claimDrawings(zip, SHEET, undefined, new Set())).toEqual([]);
	});
});

describe("cachedRead", () => {
	it("reads a part once and shares the promise and bytes", async () => {
		let calls = 0;
		const counting = {
			...zip,
			read: (name: string) => {
				calls++;
				return zip.read(name);
			},
		};
		const cache = new Map<string, Promise<Uint8Array>>();
		const first = cachedRead(counting, cache, MEDIA);
		const second = cachedRead(counting, cache, MEDIA);
		expect(second).toBe(first);
		expect(await second).toBe(await first);
		expect([...(await first)]).toEqual([...pictureBytes(1)]);
		expect(calls).toBe(1);
	});
});

describe("loadSheetImages", () => {
	const owned = [{ path: DRAWING, rels: DRAWING_RELS }];
	const claimed = new Set([DRAWING, DRAWING_RELS]);

	it("returns no pictures for a sheet that owns no drawing", async () => {
		expect(await loadSheetImages(zip, [], new Set(), new Map())).toEqual([]);
	});

	it("returns a drawing's pictures and caches only their image bytes", async () => {
		const media = new Map<string, Promise<Uint8Array>>();
		const [first] = await loadSheetImages(zip, owned, claimed, media);
		const [again] = await loadSheetImages(zip, owned, claimed, media);
		expect(first).toMatchObject({ mime: "image/png", anchor: { from: { col: 1, row: 1 } } });
		expect(again?.bytes).toBe(first?.bytes);
		expect([...media.keys()]).toEqual([MEDIA]);
	});

	it("skips a drawing without its rels part, and an image path claimed in another role", async () => {
		expect(
			await loadSheetImages(zip, [{ path: DRAWING, rels: undefined }], claimed, new Map()),
		).toEqual([]);
		expect(await loadSheetImages(zip, owned, new Set([...claimed, MEDIA]), new Map())).toEqual(
			[],
		);
	});
});
