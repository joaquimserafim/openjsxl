import { describe, expect, it } from "vitest";
import { writeXlsx } from "../../writer";
import { openZip } from "../../zip";
import { openXlsx, streamSheetRows } from "../workbook";

// F11.53: `ReadOptions.maxTotalBytes` reaches openXlsx and streamSheetRows through openZip. A
// workbook under the budget opens as before; one whose parts together pass it fails typed, even when
// every part alone fits. Absent, undefined or null means no budget.

const typed = (code: string) => expect.objectContaining({ name: "XlsxError", code });

// Three sheets of 300 rows, so a workbook is several parts of similar size.
async function workbook() {
	const rows = Array.from({ length: 300 }, (_, i) => [i, `row ${i}`, i * 2]);
	const bytes = await writeXlsx({
		sheets: ["A", "B", "C"].map((name) => ({ name, rows })),
	});
	const sizes = [...openZip(bytes).entries.values()].map((entry) => entry.uncompressedSize);
	return { bytes, largest: Math.max(...sizes), all: sizes.reduce((x, y) => x + y, 0) };
}

async function cells(bytes: Uint8Array, options?: object) {
	const book = await openXlsx(bytes, options);
	const out: unknown[] = [];
	for (const info of book.sheets) {
		for await (const row of book.sheet(info.name).rows())
			out.push(row.cells.map((c) => c.value));
	}
	return out;
}

async function streamed(bytes: Uint8Array, options?: object) {
	const out: unknown[] = [];
	for await (const row of streamSheetRows(bytes, "C", options))
		out.push(row.cells.map((c) => c.value));
	return out;
}

describe("maxTotalBytes through openXlsx and streamSheetRows", () => {
	it("leaves a workbook under the budget as it was", async () => {
		const { bytes, all } = await workbook();
		const plain = await cells(bytes);
		expect(plain).toHaveLength(900);
		expect(await cells(bytes, { maxTotalBytes: all })).toEqual(plain);
		expect(await streamed(bytes, { maxTotalBytes: all })).toEqual(await streamed(bytes));
	});

	it("fails typed when the parts together pass the budget, though each part alone fits", async () => {
		const { bytes, largest } = await workbook();
		const options = { maxTotalBytes: largest + 1 };
		await expect(cells(bytes, options)).rejects.toThrow(typed("part-too-large"));
		await expect(streamed(bytes, options)).rejects.toThrow(typed("part-too-large"));
	});

	it("takes Infinity, undefined and null as no budget, and 0 as a budget nothing fits", async () => {
		const { bytes } = await workbook();
		const plain = await cells(bytes);
		for (const maxTotalBytes of [Number.POSITIVE_INFINITY, undefined, null]) {
			expect(await cells(bytes, { maxTotalBytes })).toEqual(plain);
			expect(await streamed(bytes, { maxTotalBytes })).toHaveLength(300);
		}
		await expect(cells(bytes, { maxTotalBytes: 0 })).rejects.toThrow(typed("part-too-large"));
		await expect(streamed(bytes, { maxTotalBytes: 0 })).rejects.toThrow(
			typed("part-too-large"),
		);
	});

	it("refuses NaN, text and a negative number, typed", async () => {
		const { bytes } = await workbook();
		for (const maxTotalBytes of [Number.NaN, "abc", -1]) {
			await expect(cells(bytes, { maxTotalBytes })).rejects.toThrow(typed("invalid-input"));
			await expect(streamed(bytes, { maxTotalBytes })).rejects.toThrow(
				typed("invalid-input"),
			);
		}
	});
});
