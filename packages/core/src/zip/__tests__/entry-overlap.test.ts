import { listFixtures, loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { detectSpreadsheetFormat } from "../../reader/detect";
import { openOds } from "../../reader/ods";
import { openXlsx, streamSheetRows } from "../../reader/workbook";
import { openXlsb } from "../../reader/xlsb";
import { writeXlsx } from "../../writer";
import { openZip } from "../central-directory";
import {
	aliasEntry,
	readCentralDirectory,
	reverseCentralDirectory,
	withCompressedSize,
	withLocalHeaderOffset,
} from "./entry-overlap-fixtures";

// F11.52: no two entries may share bytes, and no entry's data may reach the central directory.
// Many central entries pointing at one local header made one small deflate stream inflate once per
// entry (100 aliases of an 8 KB stream cost about 869 MiB). openZip now rejects such an archive
// before any part is read.

const SHARE = /share bytes/;
const INTO_DIRECTORY = /runs past the start of the central directory/;

const sharedBytes = expect.objectContaining({
	name: "XlsxError",
	code: "corrupt-zip",
	message: expect.stringMatching(SHARE),
});

const twoSheets = await writeXlsx({
	sheets: [
		{ name: "A", rows: [[1]] },
		{ name: "B", rows: [[2]] },
	],
});
const aliasedXlsx = aliasEntry(twoSheets, "xl/worksheets/sheet2.xml", "xl/worksheets/sheet1.xml");
const aliasedXlsb = aliasEntry(
	await loadFixture("xlsb-basic.xlsb"),
	"xl/worksheets/sheet3.bin",
	"xl/worksheets/sheet2.bin",
);
const aliasedOds = aliasEntry(await loadFixture("odf-basic.ods"), "meta.xml", "styles.xml");

/** The archive's records sorted by local-header offset. */
const byLocalOffset = (bytes: Uint8Array) =>
	[...readCentralDirectory(bytes).records].sort(
		(a, b) => a.localHeaderOffset - b.localHeaderOffset,
	);

describe("openZip entry overlap", () => {
	it("opens the unpatched archive", () => {
		expect(() => openZip(twoSheets)).not.toThrow();
	});

	it("rejects two central entries that point at one local header", () => {
		expect(() => openZip(aliasedXlsx)).toThrow(sharedBytes);
	});

	it("rejects an entry whose data runs one byte into the next local header", () => {
		const [first, second] = byLocalOffset(twoSheets);
		if (first === undefined || second === undefined) throw new Error("expected two entries");
		const size = second.localHeaderOffset - first.dataStart + 1;
		expect(() => openZip(withCompressedSize(twoSheets, first.name, size))).toThrow(sharedBytes);
	});

	it("rejects an entry whose data runs one byte into the central directory", () => {
		const last = byLocalOffset(twoSheets).at(-1);
		if (last === undefined) throw new Error("expected an entry");
		const size = readCentralDirectory(twoSheets).offset - last.dataStart + 1;
		expect(() => openZip(withCompressedSize(twoSheets, last.name, size))).toThrow(
			expect.objectContaining({
				code: "corrupt-zip",
				message: expect.stringMatching(INTO_DIRECTORY),
			}),
		);
	});

	// Such an entry has no span to check and can never be read, so it cannot share bytes.
	it("opens an archive with an entry that has no local header, failing only on that entry's read", async () => {
		const zip = openZip(withLocalHeaderOffset(twoSheets, "xl/worksheets/sheet2.xml", 1));
		await expect(zip.read("xl/worksheets/sheet2.xml")).rejects.toEqual(
			expect.objectContaining({
				code: "corrupt-zip",
				message: expect.stringMatching(/bad local header/),
			}),
		);
		expect(await zip.read("xl/worksheets/sheet1.xml")).toEqual(
			await openZip(twoSheets).read("xl/worksheets/sheet1.xml"),
		);
	});

	it("accepts a central directory listed out of local-header order", async () => {
		const original = openZip(twoSheets);
		const reversed = openZip(reverseCentralDirectory(twoSheets));
		expect([...reversed.entries.keys()]).toEqual([...original.entries.keys()].reverse());
		for (const name of original.entries.keys()) {
			expect(await reversed.read(name), name).toEqual(await original.read(name));
		}
	});
});

describe("readers reject entries that share bytes", () => {
	it("in openXlsx and streamSheetRows", async () => {
		await expect(openXlsx(aliasedXlsx)).rejects.toEqual(sharedBytes);
		await expect(streamSheetRows(aliasedXlsx).next()).rejects.toEqual(sharedBytes);
	});

	it("in openXlsb", async () => {
		await expect(openXlsb(aliasedXlsb)).rejects.toEqual(sharedBytes);
	});

	it("in openOds", async () => {
		await expect(openOds(aliasedOds)).rejects.toEqual(sharedBytes);
	});

	it("in detectSpreadsheetFormat, which classifies none of them", async () => {
		for (const bytes of [aliasedXlsx, aliasedXlsb, aliasedOds]) {
			expect(await detectSpreadsheetFormat(bytes)).toBeUndefined();
		}
	});
});

describe("committed corpus", () => {
	it("has no archive whose entries share bytes or reach the central directory", async () => {
		let archives = 0;
		for (const name of listFixtures()) {
			const bytes = await loadFixture(name);
			if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) continue; // "PK": a ZIP container
			archives++;
			expect(() => openZip(bytes), name).not.toThrow(SHARE);
			expect(() => openZip(bytes), name).not.toThrow(INTO_DIRECTORY);
		}
		expect(archives).toBeGreaterThanOrEqual(50);
	});
});
