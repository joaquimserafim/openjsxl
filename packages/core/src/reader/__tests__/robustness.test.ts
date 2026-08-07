import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx, streamSheetRows } from "../workbook";

function corruptCentralCrc(bytes: Uint8Array, target: string): Uint8Array {
	const copy = bytes.slice();
	const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
	const eocd = copy.byteLength - 22; // committed fixtures carry no ZIP comment
	const entryCount = view.getUint16(eocd + 10, true);
	let pos = view.getUint32(eocd + 16, true);
	const decoder = new TextDecoder();
	for (let i = 0; i < entryCount; i++) {
		const nameLength = view.getUint16(pos + 28, true);
		const extraLength = view.getUint16(pos + 30, true);
		const commentLength = view.getUint16(pos + 32, true);
		const name = decoder.decode(copy.subarray(pos + 46, pos + 46 + nameLength));
		if (name === target) {
			view.setUint32(pos + 16, view.getUint32(pos + 16, true) ^ 1, true);
			return copy;
		}
		pos += 46 + nameLength + extraLength + commentLength;
	}
	throw new Error(`fixture has no central-directory entry named ${target}`);
}

async function consume(iterable: AsyncIterable<unknown>): Promise<void> {
	for await (const _value of iterable) {
		// Reaching natural EOF is what completes streaming size/CRC verification.
	}
}

// Missing optional parts must degrade gracefully, never throw (F2.4b). minimal.xlsx is a
// generated workbook with numbers and a boolean only, so it carries neither sharedStrings.xml
// nor styles.xml.

describe("reader — missing optional parts", () => {
	it("opens a workbook that has no styles.xml or sharedStrings.xml", async () => {
		const wb = await openXlsx(await loadFixture("minimal.xlsx"));
		expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
		const sheet = wb.sheet("Sheet1");
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "number", value: 1 });
		expect(sheet.cell("B1")).toEqual({ ref: "B1", type: "boolean", value: true });
		// No style table ⇒ no resolvable number format, but no error.
		expect(sheet.numberFormat("A1")).toBeUndefined();
		// Metadata accessors are empty, not throwing.
		expect(sheet.mergedCells).toEqual([]);
		expect(sheet.hyperlinks).toEqual([]);
		expect(sheet.comments).toEqual([]);
	});

	it("streams rows from a workbook with no optional parts", async () => {
		const bytes = await loadFixture("minimal.xlsx");
		const rows = [];
		for await (const row of streamSheetRows(bytes)) rows.push(row.cells.length);
		expect(rows).toEqual([2]);
	});
});

describe("reader — maxPartBytes (zip-bomb guard)", () => {
	it("rejects a part larger than the configured limit", async () => {
		const bytes = await loadFixture("basic.xlsx");
		await expect(openXlsx(bytes, { maxPartBytes: 1 })).rejects.toMatchObject({
			code: "part-too-large",
		});
	});

	it("opens normally when the limit is generous", async () => {
		const bytes = await loadFixture("basic.xlsx");
		const wb = await openXlsx(bytes, { maxPartBytes: 1_000_000 });
		expect(wb.sheet("Sheet1").cell("A1")).toEqual({
			ref: "A1",
			type: "string",
			value: "hello",
		});
	});

	it("applies the limit to streaming too", async () => {
		const bytes = await loadFixture("basic.xlsx");
		await expect(
			streamSheetRows(bytes, undefined, { maxPartBytes: 1 }).next(),
		).rejects.toMatchObject({ code: "part-too-large" });
	});
});

describe("reader — optional ZIP checksum verification", () => {
	it("threads verifyChecksums through eager and streaming reads", async () => {
		const damaged = corruptCentralCrc(
			await loadFixture("basic.xlsx"),
			"xl/worksheets/sheet1.xml",
		);

		// The 1.0.1 default remains tolerant, while opting in rejects the same worksheet through
		// both the eager zip.read() path and streamSheetRows' actual zip.readStream() path.
		expect((await openXlsx(damaged)).sheet("Sheet1").cell("A1").value).toBe("hello");
		await expect(openXlsx(damaged, { verifyChecksums: true })).rejects.toMatchObject({
			code: "corrupt-zip",
		});
		await expect(
			consume(streamSheetRows(damaged, undefined, { verifyChecksums: true })),
		).rejects.toMatchObject({ code: "corrupt-zip" });
	});
});
