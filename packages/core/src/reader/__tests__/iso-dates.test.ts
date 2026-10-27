import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { evaluateCell } from "../../formula";
import { dateToSerial } from "../../ooxml";
import type { Cell } from "../../types";
import { workbookToInput, writeXlsx } from "../../writer";
import { openXlsx, streamSheetRows } from "../workbook";
import { rewriteParts } from "./shared-part-fixtures";

// F11.60: openpyxl's Workbook(iso_dates=True) writes dates as t="d" ISO 8601 text. The buffered and
// streamed readers, the evaluator and the bridge must all see them as dates. A bare time ("13:45:30",
// A5) is outside the scoped grammar and reads as empty.

const FIXTURE = "openpyxl-iso-dates.xlsx";
const utc = (...parts: [number, number, number, number?, number?, number?, number?]) => {
	const [y, m, d, h = 0, mi = 0, s = 0, ms = 0] = parts;
	const date = new Date(Date.UTC(y, m - 1, d, h, mi, s, ms));
	if (y < 100) date.setUTCFullYear(y);
	return date;
};
const expected: Record<string, Cell> = {
	A1: { ref: "A1", type: "date", value: utc(2024, 3, 15) },
	A2: { ref: "A2", type: "date", value: utc(2024, 3, 15, 13, 45, 30) },
	A3: { ref: "A3", type: "date", value: utc(2024, 3, 15, 13, 45, 30, 123) },
	A4: { ref: "A4", type: "date", value: utc(2024, 2, 29) },
	A5: { ref: "A5", type: "empty", value: null },
	A6: { ref: "A6", type: "date", value: utc(1, 1, 1, 6) },
	A7: { ref: "A7", type: "date", value: utc(9999, 12, 31, 23, 59, 59) },
	C1: { ref: "C1", type: "string", value: "label" },
	C2: { ref: "C2", type: "number", value: 42 },
};

describe('openpyxl iso_dates (t="d") cells', () => {
	it("read as dates through cell()", async () => {
		const sheet = (await openXlsx(await loadFixture(FIXTURE))).sheet("Dates");
		for (const [ref, cell] of Object.entries(expected))
			expect(sheet.cell(ref), ref).toEqual(cell);
		expect([sheet.numberFormat("A1"), sheet.numberFormat("A4")]).toEqual([
			"yyyy-mm-dd h:mm:ss",
			"yyyy-mm-dd",
		]);
	});

	it("read the same through rows() and streamSheetRows", async () => {
		const bytes = await loadFixture(FIXTURE);
		const sheet = (await openXlsx(bytes)).sheet("Dates");
		const buffered: Cell[] = [];
		for await (const row of sheet.rows()) buffered.push(...row.cells);
		const streamed: Cell[] = [];
		for await (const row of streamSheetRows(bytes, "Dates")) streamed.push(...row.cells);
		expect(streamed).toEqual(buffered);
		for (const [ref, cell] of Object.entries(expected)) {
			if (cell.type !== "empty")
				expect(
					buffered.find((c) => c.ref === ref),
					ref,
				).toEqual(cell);
		}
	});

	it("evaluate as their serial: =A1+1", async () => {
		const wb = await openXlsx(await loadFixture(FIXTURE));
		expect(await evaluateCell(wb, "Dates", "B1")).toBe(dateToSerial(utc(2024, 3, 15)) + 1);
		expect(await evaluateCell(wb, "Dates", "B1")).toBe(45_367);
	});

	it("survive the bridge as dates with their number formats", async () => {
		const wb = await openXlsx(await loadFixture(FIXTURE));
		const sheet = (await openXlsx(await writeXlsx(await workbookToInput(wb)))).sheet("Dates");
		for (const [ref, cell] of Object.entries(expected))
			expect(sheet.cell(ref), ref).toEqual(cell);
		expect([sheet.numberFormat("A2"), sheet.numberFormat("A4")]).toEqual([
			"yyyy-mm-dd h:mm:ss",
			"yyyy-mm-dd",
		]);
	});

	it("bridge under a non-date format as their serial, and gain a date format when unstyled", async () => {
		// Recorded behavior, pending an owner ruling: the writer keeps a cell's own number format,
		// so a t="d" date under 0.00 is written as its serial under 0.00 (what Excel shows for the
		// original too) and reads back as a number. An unstyled date gains the writer's mm-dd-yy.
		const bytes = await rewriteParts(await loadFixture(FIXTURE), {
			"xl/styles.xml": (t) =>
				t
					.replace('<cellXfs count="4">', '<cellXfs count="5">')
					.replace(
						"</cellXfs>",
						'<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" /></cellXfs>',
					),
			"xl/worksheets/sheet1.xml": (t) =>
				t
					.replace('<c r="A1" s="1" t="d">', '<c r="A1" t="d">')
					.replace('<c r="A2" s="1" t="d">', '<c r="A2" s="4" t="d">'),
		});
		const source = (await openXlsx(bytes)).sheet("Dates");
		expect([source.cell("A1"), source.cell("A2")]).toEqual([expected.A1, expected.A2]);
		expect([source.style("A1"), source.numberFormat("A2")]).toEqual([undefined, "0.00"]);
		const bridged = (
			await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(bytes))))
		).sheet("Dates");
		expect(bridged.cell("A2")).toEqual({
			ref: "A2",
			type: "number",
			value: dateToSerial(utc(2024, 3, 15, 13, 45, 30)),
		});
		expect(bridged.numberFormat("A2")).toBe("0.00");
		expect(bridged.cell("A1")).toEqual(expected.A1);
		expect(bridged.style("A1")?.numberFormat).toBe("mm-dd-yy");
	});
});
