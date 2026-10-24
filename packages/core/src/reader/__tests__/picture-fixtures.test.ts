import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";
import { DRAWING_REL, legacySheetImages, pictureBytes, pictureWorkbook } from "./picture-fixtures";

describe("picture fixtures", () => {
	it("makes distinct picture payloads", () => {
		expect([...pictureBytes(1)]).toEqual([0x89, 0x50, 0x4e, 0x47, 1]);
		expect(pictureBytes(1)).not.toEqual(pictureBytes(2));
	});

	it("builds a workbook with one picture on each sheet that has bytes", async () => {
		const book = await openXlsx(await pictureWorkbook([pictureBytes(1), undefined]));
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "S2"]);
		expect((await book.sheet("S1").images()).map((i) => [...i.bytes])).toEqual([
			[...pictureBytes(1)],
		]);
		expect(await book.sheet("S2").images()).toEqual([]);
	});

	it("loads a sheet's pictures the way 1.0.0 did", async () => {
		const bytes = await pictureWorkbook([pictureBytes(3)]);
		const [picture] = await legacySheetImages(bytes, "xl/worksheets/sheet1.xml");
		expect(picture).toMatchObject({ mime: "image/png", anchor: { from: { col: 1, row: 1 } } });
		expect([...(picture?.bytes ?? [])]).toEqual([...pictureBytes(3)]);
		expect(await legacySheetImages(bytes, "xl/worksheets/no-such-sheet.xml")).toEqual([]);
	});

	it("names the OPC drawing relationship type", () => {
		expect(DRAWING_REL.endsWith("/relationships/drawing")).toBe(true);
	});
});
