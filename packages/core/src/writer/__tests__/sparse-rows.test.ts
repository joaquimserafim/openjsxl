import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { openXlsx } from "../../reader/workbook";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { CellInput } from "../types";
import { writeXlsx } from "../workbook";
import { countingRow, farRightRows, type RowReads, sheetXml } from "./sparse-row-fixtures";

// F11.54: write cost follows populated cells, not row length. Both writers used to visit every index
// of a row, so one cell at column XFD cost 16,384 visits; the bridge builds exactly such holey rows.

describe("sparse rows in the buffered writer", () => {
	it("probes a far-right row a bounded number of times, not once per column", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const xml = sheetXml({
			name: "S",
			rows: [countingRow(farRightRows(1, 16_384)[0] ?? [], reads)],
		});
		expect(xml).toContain('<row r="1"><c r="XFD1"><v>0</v></c></row>');
		expect(reads.probes).toBeLessThan(100);
		expect(reads.keyListings).toBe(1);
	});

	it("is linear in populated cells on far-right rows", () => {
		const ratio = bestDoublingRatio((n) => {
			const rows = farRightRows(n, n);
			return () => sheetXml({ name: "S", rows });
		}, 2_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);

	it("writes a holey row exactly like the same row filled with undefined and null", () => {
		const holey: CellInput[] = [];
		holey[0] = 1;
		holey[700] = "x";
		holey[16_383] = true;
		const filled: CellInput[] = Array.from({ length: 16_384 }, (_, c) =>
			c % 2 === 0 ? undefined : null,
		);
		filled[0] = 1;
		filled[700] = "x";
		filled[16_383] = true;
		const expected =
			'<row r="1"><c r="A1"><v>1</v></c><c r="ZY1" t="inlineStr"><is><t>x</t></is></c><c r="XFD1" t="b"><v>1</v></c></row>';
		expect(sheetXml({ name: "S", rows: [holey] })).toContain(expected);
		expect(sheetXml({ name: "S", rows: [filled] })).toContain(expected);
	});

	it("lists a reused sparse row object's keys once per sheet, however many named keys it has", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const base = farRightRows(1, 16_384)[0] ?? [];
		for (let k = 0; k < 1000; k++) Object.defineProperty(base, `p${k}`, { value: k });
		const row = countingRow(base, reads);
		const xml = sheetXml({ name: "S", rows: Array.from({ length: 50 }, () => row) });
		expect(xml).toContain('<row r="50"><c r="XFD50"><v>0</v></c></row>');
		expect(reads.keyListings).toBe(1);
	});

	it("sees a cell that an earlier cell's getter fills, as an index loop does", () => {
		const row: CellInput[] = [];
		Object.defineProperty(row, 0, {
			enumerable: true,
			get() {
				row[2] = "added";
				return "first";
			},
		});
		row.length = 3;
		expect(sheetXml({ name: "S", rows: [row] })).toContain(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c><c r="C1" t="inlineStr"><is><t>added</t></is></c></row>',
		);
	});

	it("renders a captured table header even if a later header getter deletes it", () => {
		const header: CellInput[] = ["A"];
		Object.defineProperty(header, 1, {
			enumerable: true,
			get() {
				delete header[0];
				return "B";
			},
		});
		const xml = sheetXml({
			name: "S",
			rows: [header, [1, 2]],
			tables: [{ name: "T", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false }],
		});
		expect(xml).toContain(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row>',
		);
	});

	// The row length is read once and every column comes from below it (the cell half of F11.42(a)).
	it("reads a row's length once, so a length that grows on a later read adds no cells", () => {
		const target: CellInput[] = ["a"];
		target[20_000] = "far";
		let reads = 0;
		const row = new Proxy(target, {
			get(t, key, receiver) {
				if (key === "length") return ++reads === 1 ? 1 : 20_001;
				return Reflect.get(t, key, receiver);
			},
		});
		const xml = sheetXml({ name: "S", rows: [row] });
		expect(xml).toContain('<dimension ref="A1"/>');
		expect(xml).not.toContain("ACOG1");
		expect(reads).toBe(1);
	});

	it("still writes a styled empty cell far to the right", () => {
		const row: CellInput[] = [];
		row[5000] = { value: null, style: { font: { bold: true } } };
		expect(sheetXml({ name: "S", rows: [row] })).toContain('<c r="GJI1" s="1"/>');
	});

	it("orders cells by column whatever order a Proxy row lists its keys in", () => {
		const row: CellInput[] = [];
		row[0] = "a";
		row[500] = "b";
		row[999] = "c";
		const reads: RowReads = { probes: 0, keyListings: 0 };
		expect(sheetXml({ name: "S", rows: [countingRow(row, reads, true)] })).toBe(
			sheetXml({ name: "S", rows: [row] }),
		);
		expect(reads.keyListings).toBe(1);
	});

	it("writes a non-enumerable own cell and skips an inherited one", () => {
		const row: CellInput[] = [];
		row[0] = 1;
		Object.defineProperty(row, 900, {
			value: 5,
			enumerable: false,
			writable: true,
			configurable: true,
		});
		row[999] = 2;
		const proto: object = Object.create(Array.prototype, {
			500: { value: "inherited", enumerable: true },
		});
		Object.setPrototypeOf(row, proto);
		expect(Array.isArray(row)).toBe(true);
		expect(sheetXml({ name: "S", rows: [row] })).toContain(
			'<row r="1"><c r="A1"><v>1</v></c><c r="AHQ1"><v>5</v></c><c r="ALL1"><v>2</v></c></row>',
		);
	});

	it("names a table from header cells in a sparse header row", () => {
		const header: CellInput[] = [];
		header[99] = "Region";
		header[100] = "Total";
		const data: CellInput[] = [];
		data[99] = "West";
		data[100] = 10;
		const xml = sheetXml({
			name: "S",
			rows: [header, data],
			tables: [{ name: "T", ref: "CV1:CW2", columns: [], headerRow: true, totalsRow: false }],
		});
		expect(xml).toContain(
			'<row r="1"><c r="CV1" t="inlineStr"><is><t>Region</t></is></c><c r="CW1" t="inlineStr"><is><t>Total</t></is></c></row>',
		);
	});
});

describe("sparse rows across writers and the bridge", () => {
	it("reads back the same far-right cells from both writers and after the bridge", async () => {
		const rows = farRightRows(3, 16_384);
		const buffered = await writeXlsx({ sheets: [{ name: "S", rows }] });
		const streamed = new Uint8Array(
			await new Response(streamXlsx({ sheets: [{ name: "S", rows }] })).arrayBuffer(),
		);
		const bridged = await writeXlsx(await workbookToInput(await openXlsx(buffered)));
		for (const bytes of [buffered, streamed, bridged]) {
			const sheet = (await openXlsx(bytes)).sheet("S");
			expect(["XFD1", "XFD2", "XFD3", "XFC1"].map((ref) => sheet.cell(ref).value)).toEqual([
				0,
				1,
				2,
				null,
			]);
		}
		expect(bridged).toEqual(buffered);
	});

	it("lists a reused sparse row object's keys once per streamed sheet", async () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const row = countingRow(farRightRows(1, 16_384)[0] ?? [], reads);
		const rows = Array.from({ length: 50 }, () => row);
		const streamed = new Uint8Array(
			await new Response(streamXlsx({ sheets: [{ name: "S", rows }] })).arrayBuffer(),
		);
		expect((await openXlsx(streamed)).sheet("S").cell("XFD50").value).toBe(0);
		expect(reads.keyListings).toBe(1);
	});
});
