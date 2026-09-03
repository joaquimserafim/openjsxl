import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { writeXlsx } from "../../writer/workbook";
import { openCsv } from "../csv";
import { openOds } from "../ods";
import { openXlsx } from "../workbook";
import { openXlsb } from "../xlsb";

// F11.10 — `Workbook.date1904` surfaces the workbook's date system on the shared model so the formula
// engine (and any consumer converting serials) uses the right epoch. XLSX reports `<workbookPr
// date1904>`; ODS/CSV have no such notion (false); XLSB reports the parsed `BrtWbProp` flag, which
// F11.13 will read — until then the reader's documented 1900 default (false).

describe("Workbook.date1904", () => {
	it("is false for a default-epoch .xlsx and true when written with { date1904: true }", async () => {
		const rows = [[new Date(Date.UTC(1904, 0, 1))]];
		const plain = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		expect(plain.date1904).toBe(false);
		const mac = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows }] }, { date1904: true }),
		);
		expect(mac.date1904).toBe(true);
		// The flag and the cell typing agree: both files read the same calendar date.
		expect(plain.sheet("S").cell("A1").value).toEqual(new Date(Date.UTC(1904, 0, 1)));
		expect(mac.sheet("S").cell("A1").value).toEqual(new Date(Date.UTC(1904, 0, 1)));
	});

	it("reads the real-producer 1904 fixture as true", async () => {
		expect((await openXlsx(await loadFixture("date_1904.xlsx"))).date1904).toBe(true);
	});

	it("is false for .ods and .csv, and false for .xlsb until F11.13 parses BrtWbProp", async () => {
		expect((await openOds(await loadFixture("odf-basic.ods"))).date1904).toBe(false);
		expect(openCsv("a,b\n1,2\n").date1904).toBe(false);
		expect((await openXlsb(await loadFixture("xlsb-basic.xlsb"))).date1904).toBe(false);
	});
});
