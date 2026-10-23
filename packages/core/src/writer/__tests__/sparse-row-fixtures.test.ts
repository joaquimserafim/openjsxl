import { describe, expect, it } from "vitest";
import { countingRow, farRightRows, type RowReads, sheetXml } from "./sparse-row-fixtures";

describe("sparse-row fixtures", () => {
	it("builds holey rows with one far-right value each", () => {
		const rows = farRightRows(2, 3);
		expect(rows.map((row) => row.length)).toEqual([3, 3]);
		expect(rows.map((row) => Object.keys(row))).toEqual([["2"], ["2"]]);
		expect(rows.map((row) => row[2])).toEqual([0, 1]);
	});

	it("renders a sheet through the buffered writer", () => {
		expect(sheetXml({ name: "S", rows: [[1]] })).toContain(
			'<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
		);
	});

	it("counts own-index probes and key listings, optionally reversing the keys", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const row = countingRow([1, 2], reads, true);
		expect([Object.hasOwn(row, 0), Object.hasOwn(row, 5)]).toEqual([true, false]);
		expect(Object.getOwnPropertyNames(row)).toEqual(["length", "1", "0"]);
		expect(reads).toEqual({ probes: 2, keyListings: 1 });
		expect(Array.isArray(row)).toBe(true);
	});
});
