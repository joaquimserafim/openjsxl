import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { RecordData, readRecords } from "../../biff";
import type { Cell, Worksheet } from "../../types";
import * as R from "../../xlsb/records";
import { openZip } from "../../zip";
import { openXlsx } from "../workbook";
import { openXlsb } from "../xlsb";

// F11.37: SheetJS 0.20.3 writes a width-only or hidden-only column as a BrtColInfo at XF 0, with an
// all-zero BrtWsDim for an A1-only or empty sheet. The same authored workbook written as xlsx
// declares `<dimension ref>`. Provenance in fixtures/data/README.md.
const SHEETS = ["WidthD", "Empty", "HiddenC", "Away"];

async function cells(sheet: Worksheet): Promise<Cell[]> {
	const out: Cell[] = [];
	for await (const row of sheet.rows()) out.push(...row.cells);
	return out;
}

// The four u32 fields at the start of a record: a BrtWsDim's bounds, a BrtColInfo's first, last,
// width and XF.
function fields(data: Uint8Array | undefined): number[] {
	if (data === undefined) return [];
	const d = new RecordData(data);
	return [d.u32(), d.u32(), d.u32(), d.u32()];
}

// A BrtColInfo's u16 flags after its XF: fHidden is 0x01, fUserSet (a set width) 0x02.
function columnFlags(data: Uint8Array | undefined): number {
	if (data === undefined) return -1;
	const d = new RecordData(data);
	d.skip(16);
	return d.u16();
}

describe("XLSB width-only and hidden-only columns — SheetJS fixture (F11.37)", () => {
	it("holds the producer's all-zero dimensions and XF-0 column records", async () => {
		const zip = openZip(await loadFixture("sheetjs-width-only.xlsb"), {
			verifyChecksums: true,
		});
		const expected = [
			{ bounds: [0, 0, 0, 0], column: [3, 3], flags: 0x02 },
			{ bounds: [0, 0, 0, 0], column: [1, 1], flags: 0x02 },
			{ bounds: [0, 0, 0, 0], column: [2, 2], flags: 0x01 },
			{ bounds: [1, 2, 1, 2], column: [5, 5], flags: 0x02 },
		];
		for (const [index, { bounds, column, flags }] of expected.entries()) {
			const records = [...readRecords(await zip.read(`xl/worksheets/sheet${index + 1}.bin`))];
			expect(fields(records.find((record) => record.id === R.DIMENSION)?.data)).toEqual(
				bounds,
			);
			const columns = records.filter((record) => record.id === R.COL_INFO);
			expect(columns).toHaveLength(1);
			const [first, last, , ixfe] = fields(columns[0]?.data);
			expect([first, last, ixfe]).toEqual([...column, 0]);
			expect(columns[0]?.data.length).toBe(18);
			expect(columnFlags(columns[0]?.data)).toBe(flags);
		}
	});

	it("reads each sheet's dimension without its width-only or hidden-only column", async () => {
		const book = await openXlsb(await loadFixture("sheetjs-width-only.xlsb"));
		expect(SHEETS.map((name) => book.sheet(name).dimension)).toEqual([
			"A1",
			undefined,
			"A1",
			"B2:C3",
		]);
	});

	it("agrees with the same workbook's xlsx dimension wherever both declare cells", async () => {
		const xlsb = await openXlsb(await loadFixture("sheetjs-width-only.xlsb"));
		const xlsx = await openXlsx(await loadFixture("sheetjs-width-only.xlsx"));
		const compared: string[] = [];
		for (const name of SHEETS) {
			const xlsbCells = await cells(xlsb.sheet(name));
			expect(xlsbCells).toEqual(await cells(xlsx.sheet(name)));
			if (xlsbCells.length === 0) continue;
			expect(xlsb.sheet(name).dimension, name).toBe(xlsx.sheet(name).dimension);
			compared.push(name);
		}
		expect(compared).toEqual(["WidthD", "HiddenC", "Away"]);
	});
});
