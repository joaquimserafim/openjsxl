import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openCsv } from "../csv";
import { openXlsx } from "../workbook";
import { parsePrintSetup } from "../worksheet";

// F10.4 — print setup: printOptions, pageMargins, pageSetup, headerFooter. The reader keeps only
// explicitly-present values, clamps/drops hostile numerics, and reads only a DIRECT <worksheet> child.

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet ${NS}><sheetData/>${body}</worksheet>`;

describe("parsePrintSetup", () => {
	it("reads printOptions booleans (only present ones)", () => {
		expect(
			parsePrintSetup(sheet('<printOptions horizontalCentered="1" gridLines="1"/>'))
				.printOptions,
		).toEqual({ horizontalCentered: true, gridLines: true });
	});

	it("reads all six page margins", () => {
		const body =
			'<pageMargins left="0.7" right="0.7" top="1" bottom="1" header="0.3" footer="0.3"/>';
		expect(parsePrintSetup(sheet(body)).pageMargins).toEqual({
			left: 0.7,
			right: 0.7,
			top: 1,
			bottom: 1,
			header: 0.3,
			footer: 0.3,
		});
	});

	it("DROPS pageMargins when a value is missing or non-finite; clamps negatives/over-max", () => {
		// Missing `footer` → the whole element drops (all six are required).
		expect(
			parsePrintSetup(
				sheet('<pageMargins left="1" right="1" top="1" bottom="1" header="1"/>'),
			).pageMargins,
		).toBeUndefined();
		// A NaN/Infinity value drops the element.
		expect(
			parsePrintSetup(
				sheet('<pageMargins left="x" right="1" top="1" bottom="1" header="1" footer="1"/>'),
			).pageMargins,
		).toBeUndefined();
		// A negative clamps to 0, an over-max clamps to 49.
		expect(
			parsePrintSetup(
				sheet(
					'<pageMargins left="-5" right="9999" top="1" bottom="1" header="1" footer="1"/>',
				),
			).pageMargins,
		).toEqual({ left: 0, right: 49, top: 1, bottom: 1, header: 1, footer: 1 });
	});

	it("reads curated pageSetup attrs; clamps scale to 10..400; drops a hostile uint", () => {
		expect(
			parsePrintSetup(
				sheet(
					'<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/>',
				),
			).pageSetup,
		).toEqual({ orientation: "landscape", paperSize: 9, fitToWidth: 1, fitToHeight: 0 });
		expect(parsePrintSetup(sheet('<pageSetup scale="5"/>')).pageSetup).toEqual({ scale: 10 });
		expect(parsePrintSetup(sheet('<pageSetup scale="900"/>')).pageSetup).toEqual({
			scale: 400,
		});
		// A 21-digit paperSize is dropped (not coerced to a lossy float), and unknown orientation ignored.
		expect(
			parsePrintSetup(
				sheet('<pageSetup paperSize="999999999999999999999" orientation="sideways"/>'),
			).pageSetup,
		).toBeUndefined();
	});

	it("reads headerFooter child strings (with & codes) and flag attrs", () => {
		const body =
			'<headerFooter differentFirst="1"><oddHeader>&amp;CTitle</oddHeader>' +
			"<oddFooter>&amp;RPage &amp;P</oddFooter><firstHeader>&amp;CCover</firstHeader></headerFooter>";
		expect(parsePrintSetup(sheet(body)).headerFooter).toEqual({
			differentFirst: true,
			oddHeader: "&CTitle",
			oddFooter: "&RPage &P",
			firstHeader: "&CCover",
		});
	});

	it("returns an empty object when there is no print setup", () => {
		expect(parsePrintSetup(sheet(""))).toEqual({});
	});

	// Adversarial-review class (F10.2/F10.3 precedent): a <customSheetView> carries its OWN copies of all
	// four print elements; only a DIRECT <worksheet> child is the active one.
	it("ignores print elements nested inside <customSheetViews>", () => {
		const nested =
			'<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}">' +
			'<pageMargins left="9" right="9" top="9" bottom="9" header="9" footer="9"/>' +
			'<pageSetup orientation="portrait"/><printOptions gridLines="1"/>' +
			"<headerFooter><oddHeader>&amp;Cnope</oddHeader></headerFooter></customSheetView></customSheetViews>";
		expect(parsePrintSetup(sheet(nested))).toEqual({});
	});

	it("still reads the sheet-level elements when a custom view ALSO carries them", () => {
		const body =
			'<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}">' +
			'<pageSetup orientation="portrait"/></customSheetView></customSheetViews>' +
			'<pageSetup orientation="landscape"/>';
		expect(parsePrintSetup(sheet(body)).pageSetup).toEqual({ orientation: "landscape" });
	});
});

describe("reader — openpyxl-authored print-setup fixture (F10.4)", () => {
	it("reads margins, landscape+fit-to-page, print options, and header/footer", async () => {
		const s = (await openXlsx(await loadFixture("openpyxl-print-setup.xlsx"))).sheet("Report");
		expect(s.pageSetup).toEqual({
			orientation: "landscape",
			paperSize: 9,
			fitToWidth: 1,
			fitToHeight: 0,
		});
		expect(s.pageMargins?.left).toBe(0.7);
		expect(s.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
		expect(s.headerFooter?.oddHeader).toBe("&CQuarterly Report");
	});
});

describe("reader — degrade for non-xlsx formats", () => {
	it("openCsv exposes the four print accessors as undefined", () => {
		const wb = openCsv(new TextEncoder().encode("a,b\n1,2\n"));
		const s = wb.sheet(wb.sheets[0]?.name ?? "");
		expect(s.pageMargins).toBeUndefined();
		expect(s.pageSetup).toBeUndefined();
		expect(s.printOptions).toBeUndefined();
		expect(s.headerFooter).toBeUndefined();
	});
});
