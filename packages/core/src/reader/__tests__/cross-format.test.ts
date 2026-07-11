import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openCsv } from "../csv";
import { detectSpreadsheetFormat } from "../detect";
import { openOds } from "../ods";
import { openXlsx, type Workbook } from "../workbook";
import { openXlsb } from "../xlsb";

// F7.4 — the cross-format equivalence corpus. One logical table, authored identically as
// equiv.{xlsx,xlsb,ods,csv}, must read to the SAME typed value snapshot through all four readers.
// This is the guarantee behind "a user uploaded a spreadsheet" being one code path: past the reader,
// the dialect is invisible. Dates are deliberately excluded — CSV infers only numbers/booleans, so a
// date column would diverge (the documented CSV type-inference boundary).

// Snapshot EVERY populated cell of the first sheet into a ref → {type,value} map — iterating rows()
// (not a fixed ref list) so an EXTRA or out-of-range cell in any lane shows up as an extra key and
// diverges (the sheet NAME differs — CSV synthesizes "Sheet1" — so we work off sheet[0], not names).
async function valueSnapshot(
	wb: Workbook,
): Promise<Record<string, { type: string; value: unknown }>> {
	const sheet = wb.sheet(wb.sheets[0]?.name ?? "");
	const out: Record<string, { type: string; value: unknown }> = {};
	for await (const row of sheet.rows()) {
		for (const cell of row.cells) {
			out[cell.ref] = {
				type: cell.type,
				value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
			};
		}
	}
	return out;
}

const EXPECTED: Record<string, { type: string; value: unknown }> = {
	A1: { type: "string", value: "name" },
	B1: { type: "string", value: "qty" },
	C1: { type: "string", value: "active" },
	A2: { type: "string", value: "Apples" },
	B2: { type: "number", value: 42 },
	C2: { type: "boolean", value: true },
	A3: { type: "string", value: "Pears" },
	B3: { type: "number", value: 7 },
	C3: { type: "boolean", value: false },
};

describe("cross-format equivalence (F7.4)", () => {
	it("reads the same full-cell value snapshot from equiv.{xlsx,xlsb,ods,csv}", async () => {
		const xlsx = await valueSnapshot(await openXlsx(await loadFixture("equiv.xlsx")));
		const xlsb = await valueSnapshot(await openXlsb(await loadFixture("equiv.xlsb")));
		const ods = await valueSnapshot(await openOds(await loadFixture("equiv.ods")));
		const csv = await valueSnapshot(openCsv(await loadFixture("equiv.csv")));

		// Exact map equality: the same nine cells, no extras — an out-of-range cell in any lane diverges.
		expect(xlsx).toEqual(EXPECTED);
		expect(xlsb).toEqual(xlsx);
		expect(ods).toEqual(xlsx);
		expect(csv).toEqual(xlsx);
	});
});

describe("macro-enabled / template workbooks open as xlsx (F7.4)", () => {
	it("openXlsx reads a .xlsm, ignoring the content-type label and the vbaProject part", async () => {
		const bytes = await loadFixture("xlsm-macro.xlsm");
		// It classifies as xlsx (the .xlsm/.xltx family is read by openXlsx)...
		expect(await detectSpreadsheetFormat(bytes)).toBe("xlsx");
		// ...and opens like any workbook: parts resolve through the rels graph, the VBA blob is never read.
		const sheet = (await openXlsx(bytes)).sheet("Macros");
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "number", value: 42 });
		expect(sheet.cell("B1")).toEqual({ ref: "B1", type: "boolean", value: true });
	});
});
