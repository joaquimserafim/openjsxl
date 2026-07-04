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
