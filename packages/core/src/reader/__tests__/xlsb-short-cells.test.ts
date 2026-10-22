import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { readRecords } from "../../biff";
import type { Cell } from "../../types";
import { streamXlsx, workbookToInput, writeXlsx } from "../../writer";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";
import { openXlsb } from "../xlsb";

const files = ["sheetjs-short-cells-inline.xlsb", "sheetjs-short-cells-sst.xlsb"];

function grid(offset: number): Cell[] {
	return Array.from({ length: 20 }, (_, index) => ({
		ref: `${String.fromCharCode(65 + (index % 10))}${Math.floor(index / 10) + 1}`,
		type: "number",
		value: index + 1 + offset,
	}));
}

const mixed: readonly Cell[] = [
	{ ref: "A1", type: "number", value: 1 },
	{ ref: "B1", type: "string", value: "inline α😀" },
	{ ref: "C1", type: "boolean", value: true },
	{ ref: "D1", type: "boolean", value: false },
	{ ref: "E1", type: "error", value: "#DIV/0!" },
	{ ref: "F1", type: "date", value: new Date(Date.UTC(1904, 0, 1)) },
	{ ref: "G1", type: "date", value: new Date(Date.UTC(1904, 0, 2, 12)) },
	{ ref: "I1", type: "number", value: 99.125 },
	{ ref: "J1", type: "string", value: "" },
	{ ref: "A2", type: "number", value: 42 },
	{ ref: "C2", type: "string", value: "after gap" },
	{ ref: "D2", type: "boolean", value: true },
	{ ref: "G2", type: "error", value: "#N/A" },
	{ ref: "H2", type: "number", value: -7 },
	{ ref: "I2", type: "number", value: 12.5 },
	{ ref: "J2", type: "string", value: "tail" },
	{ ref: "D3", type: "number", value: 31 },
	{ ref: "E3", type: "string", value: "reset" },
	{ ref: "F3", type: "boolean", value: false },
];

async function expectValues(workbook: Workbook, bridged = false): Promise<void> {
	expect(workbook.date1904).toBe(true);
	expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Integers", "Decimals", "Mixed"]);
	const sheets = [
		{ name: "Integers", cells: grid(0) },
		{ name: "Decimals", cells: grid(0.25) },
		{ name: "Mixed", cells: mixed },
	];
	for (const { name, cells } of sheets) {
		const sheet = workbook.sheet(name);
		// Existing bridge contract: an error without formula text becomes its literal string.
		const expected = cells.map((cell) =>
			bridged && cell.type === "error" ? { ...cell, type: "string" } : cell,
		);
		for (const cell of expected)
			expect(sheet.cell(cell.ref), `${name}!${cell.ref}`).toEqual(cell);
		const iterated: Cell[] = [];
		for await (const row of sheet.rows()) iterated.push(...row.cells);
		expect(iterated, name).toEqual(expected);
	}
	for (const ref of ["H1", "B2", "E2", "F2", "A3", "B3", "C3", "G3"]) {
		expect(workbook.sheet("Mixed").cell(ref), ref).toEqual({ ref, type: "empty", value: null });
	}
}

describe.each(files)("XLSB short cells — real SheetJS producer: %s", (file) => {
	it("contains the producer's full and short records, including its selected string encoding", async () => {
		const zip = openZip(await loadFixture(file), { verifyChecksums: true });
		const integerRecords = [...readRecords(await zip.read("xl/worksheets/sheet1.bin"))];
		expect(integerRecords.filter((record) => record.id === 0x02)).toHaveLength(2);
		expect(integerRecords.filter((record) => record.id === 0x0d)).toHaveLength(18);
		const decimalRecords = [...readRecords(await zip.read("xl/worksheets/sheet2.bin"))];
		expect(decimalRecords.filter((record) => record.id === 0x05)).toHaveLength(2);
		expect(decimalRecords.filter((record) => record.id === 0x10)).toHaveLength(18);
		const mixedRecords = [...readRecords(await zip.read("xl/worksheets/sheet3.bin"))];
		const ids = new Set(mixedRecords.map((record) => record.id));
		for (const id of [0x01, 0x0c, 0x0d, 0x0e, 0x0f, 0x10])
			expect(ids.has(id), String(id)).toBe(true);
		const shared = file.endsWith("-sst.xlsb");
		expect(ids.has(shared ? 0x07 : 0x06)).toBe(true);
		expect(ids.has(shared ? 0x12 : 0x11)).toBe(true);
		expect(ids.has(shared ? 0x11 : 0x12)).toBe(false);
		expect(zip.has("xl/sharedStrings.bin")).toBe(shared);
	});

	it("reads every authored coordinate and iterates all values across gaps and row resets", async () => {
		const workbook = await openXlsb(await loadFixture(file));
		await expectValues(workbook);
		expect(workbook.sheet("Integers").dimension).toBe("A1:J2");
		expect(workbook.sheet("Decimals").dimension).toBe("A1:J2");
		expect(workbook.sheet("Mixed").dimension).toBe("A1:J3");
	});

	it("retains short-cell number formats and date detection without adding unsupported styles", async () => {
		const sheet = (await openXlsb(await loadFixture(file))).sheet("Mixed");
		expect(sheet.numberFormat("F1")).toBe("yyyy-mm-dd");
		expect(sheet.numberFormat("G1")).toBe("yyyy-mm-dd hh:mm:ss");
		expect(sheet.numberFormat("I2")).toBe("0.000");
		for (const ref of ["H1", "E2", "C3"]) expect(sheet.numberFormat(ref), ref).toBe("0.00");
		expect(sheet.style("I2")).toBeUndefined();
	});

	it.each(["buffered", "streaming"])(
		"preserves every value through the %s XLSX bridge",
		async (mode) => {
			const input = await workbookToInput(await openXlsb(await loadFixture(file)));
			const bytes =
				mode === "buffered"
					? await writeXlsx(input)
					: new Uint8Array(await new Response(streamXlsx(input)).arrayBuffer());
			await expectValues(await openXlsx(bytes), true);
		},
	);
});
