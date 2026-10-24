import { listFixtures, loadFixture } from "@openjsxl/fixtures";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { detectSpreadsheetFormat } from "../detect";
import { openXlsx } from "../workbook";
import { DRAWING_REL, legacySheetImages, pictureBytes, pictureWorkbook } from "./picture-fixtures";
import { rewriteParts } from "./shared-part-fixtures";

// F11.96: each drawing and media part is read at most once per workbook. A drawing part belongs to
// the first sheet in tab order that names it, repeat drawing relationships are skipped, and one
// workbook-wide media cache shares each media part's bytes. Reads are counted per part through a
// wrapper around openZip, as in F11.95.

const reads = vi.hoisted(() => new Map<string, number>());

vi.mock("../../zip", async (importOriginal) => {
	const actual = await importOriginal<typeof import("../../zip")>();
	return {
		...actual,
		openZip: (...args: Parameters<typeof actual.openZip>) => {
			const archive = actual.openZip(...args);
			const count = (name: string) => reads.set(name, (reads.get(name) ?? 0) + 1);
			return {
				...archive,
				read: (name: string) => {
					count(name);
					return archive.read(name);
				},
				readStream: (name: string) => {
					count(name);
					return archive.readStream(name);
				},
			};
		},
	};
});

// One sheet whose rels name its one-picture drawing twice.
const twoRelsOneDrawing = await rewriteParts(await pictureWorkbook([pictureBytes(1)]), {
	"xl/worksheets/_rels/sheet1.xml.rels": (t) =>
		t.replace(
			"</Relationships>",
			`<Relationship Id="rId9" Type="${DRAWING_REL}" Target="../drawings/drawing1.xml"/></Relationships>`,
		),
});
// S2's drawing relationship names S1's drawing part instead of its own.
const twoSheetsOneDrawing = await rewriteParts(
	await pictureWorkbook([pictureBytes(1), pictureBytes(2)]),
	{
		"xl/worksheets/_rels/sheet2.xml.rels": (t) =>
			t.replace("../drawings/drawing2.xml", "../drawings/drawing1.xml"),
	},
);
// One picture on two sheets: the writer stores one media part that two drawings use.
const sharedMedia = await pictureWorkbook([pictureBytes(1), pictureBytes(1)]);
// S1's picture names S2's drawing-rels part as its image, which S2 claimed at open.
const imageOnDrawingRels = await rewriteParts(
	await pictureWorkbook([pictureBytes(1), pictureBytes(2)]),
	{
		"xl/drawings/_rels/drawing1.xml.rels": (t) =>
			t.replace("../media/image1.png", "_rels/drawing2.xml.rels"),
	},
);
// S2 adds a drawing relationship to S1's image part; the claim at open takes that part.
const drawingOnImage = await rewriteParts(
	await pictureWorkbook([pictureBytes(1), pictureBytes(2)]),
	{
		"xl/worksheets/_rels/sheet2.xml.rels": (t) =>
			t.replace(
				"</Relationships>",
				`<Relationship Id="rId9" Type="${DRAWING_REL}" Target="../media/image1.png"/></Relationships>`,
			),
	},
);

beforeEach(() => reads.clear());

describe("drawing parts", () => {
	it("returns a drawing's pictures once when a sheet names it twice", async () => {
		const book = await openXlsx(twoRelsOneDrawing);
		expect(await book.sheet("S1").images()).toHaveLength(1);
		expect(reads.get("xl/drawings/drawing1.xml")).toBe(1);
	});

	it("gives a drawing named by two sheets to the first in tab order, whichever loads first", async () => {
		for (const order of [
			["S2", "S1"],
			["S1", "S2"],
		]) {
			const book = await openXlsx(twoSheetsOneDrawing);
			const loaded = new Map<string, number>();
			for (const name of order) loaded.set(name, (await book.sheet(name).images()).length);
			expect(Object.fromEntries(loaded), order.join(" then ")).toEqual({ S1: 1, S2: 0 });
		}
	});
});

describe("media parts", () => {
	it("shares one buffer between pictures in different drawings that use one media part", async () => {
		const book = await openXlsx(sharedMedia);
		const [first] = await book.sheet("S1").images();
		const [second] = await book.sheet("S2").images();
		expect(first?.bytes).toBe(second?.bytes);
		expect([...(first?.bytes ?? [])]).toEqual([...pictureBytes(1)]);
	});

	it("reads a shared media part once when two sheets load their pictures at the same time", async () => {
		const book = await openXlsx(sharedMedia);
		const [s1, s2] = await Promise.all([book.sheet("S1").images(), book.sheet("S2").images()]);
		expect(s1?.[0]?.bytes).toBe(s2?.[0]?.bytes);
		expect(reads.get("xl/media/image1.png")).toBe(1);
	});

	it("reads each picture part once through the bridge", async () => {
		await workbookToInput(await openXlsx(sharedMedia));
		expect(reads.get("xl/media/image1.png")).toBe(1);
		expect(reads.get("xl/drawings/drawing1.xml")).toBe(1);
		expect(reads.get("xl/drawings/drawing2.xml")).toBe(1);
		expect(reads.get("xl/drawings/_rels/drawing1.xml.rels")).toBe(1);
	});
});

// Crafted packages where one part plays two picture roles. Claims at open decide, so the result is the
// same in either call order: a drawing or drawing-rels part claimed at open is never read as an image.
describe("a part in two picture roles", () => {
	it.each([
		["an image named as another drawing's rels part", imageOnDrawingRels],
		["an image part named as a later sheet's drawing", drawingOnImage],
	])("skips %s, in either call order", async (_case, bytes) => {
		for (const order of [
			["S1", "S2"],
			["S2", "S1"],
		]) {
			const book = await openXlsx(bytes);
			const loaded = new Map<string, number>();
			for (const name of order) loaded.set(name, (await book.sheet(name).images()).length);
			expect(Object.fromEntries(loaded), order.join(" then ")).toEqual({ S1: 0, S2: 1 });
		}
	});
});

describe("read counts and the corpus", () => {
	it("reads every part at most once per workbook, pictures included", async () => {
		const inputs: [string, Uint8Array][] = [
			["two rels, one drawing", twoRelsOneDrawing],
			["two sheets, one drawing", twoSheetsOneDrawing],
			["shared media", sharedMedia],
			["image on drawing rels", imageOnDrawingRels],
			["drawing on image", drawingOnImage],
		];
		for (const name of listFixtures()) inputs.push([name, await loadFixture(name)]);
		let checked = 0;
		for (const [name, bytes] of inputs) {
			if ((await detectSpreadsheetFormat(bytes)) !== "xlsx") continue;
			reads.clear();
			const [opened] = await Promise.allSettled([openXlsx(bytes)]);
			if (opened?.status !== "fulfilled") continue;
			const book = opened.value;
			await Promise.all(book.sheets.map((s) => book.sheet(s.name).images()));
			expect(
				[...reads].filter(([, count]) => count > 1),
				name,
			).toEqual([]);
			checked++;
		}
		expect(checked).toBeGreaterThanOrEqual(35);
	});

	it("returns the same pictures as the 1.0.0 loader for every corpus workbook", async () => {
		let pictures = 0;
		for (const name of listFixtures()) {
			const bytes = await loadFixture(name);
			if ((await detectSpreadsheetFormat(bytes)) !== "xlsx") continue;
			const [opened] = await Promise.allSettled([openXlsx(bytes)]);
			if (opened?.status !== "fulfilled") continue;
			for (const info of opened.value.sheets) {
				const images = await opened.value.sheet(info.name).images();
				expect(images, `${name} ${info.name}`).toEqual(
					await legacySheetImages(bytes, info.path),
				);
				pictures += images.length;
			}
		}
		expect(pictures).toBeGreaterThan(0);
	});
});
