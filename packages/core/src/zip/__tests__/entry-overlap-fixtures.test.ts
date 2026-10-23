import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import {
	aliasEntry,
	centralRecord,
	readCentralDirectory,
	reverseCentralDirectory,
	withCompressedSize,
	withLocalHeaderOffset,
} from "./entry-overlap-fixtures";

const twoEntries = (): Promise<Uint8Array> =>
	writeZip([
		{ name: "a", data: new TextEncoder().encode("first") },
		{ name: "b", data: new TextEncoder().encode("second entry") },
	]);

describe("entry-overlap fixtures", () => {
	it("reads every central record with its local data offset", async () => {
		const zip = await twoEntries();
		const { offset, records } = readCentralDirectory(zip);
		expect(records.map((r) => r.name)).toEqual(["a", "b"]);
		const [a, b] = records;
		expect(a?.localHeaderOffset).toBe(0);
		expect(a?.dataStart).toBe(31); // 30-byte local header + the one-byte name
		expect(b?.localHeaderOffset).toBe((a?.dataStart ?? 0) + (a?.compressedSize ?? 0));
		expect(offset).toBe((b?.dataStart ?? 0) + (b?.compressedSize ?? 0));
	});

	it("fails without an end-of-central-directory record or a named record", async () => {
		const zip = await twoEntries();
		expect(() => readCentralDirectory(new Uint8Array(4))).toThrow(/end-of-central-directory/);
		expect(() => centralRecord(zip, "missing")).toThrow(/missing/);
	});

	it("points an alias at the source's local header with the source's sizes", async () => {
		const zip = await twoEntries();
		const aliased = aliasEntry(zip, "b", "a");
		expect(centralRecord(aliased, "b")).toMatchObject({
			localHeaderOffset: centralRecord(zip, "a").localHeaderOffset,
			compressedSize: centralRecord(zip, "a").compressedSize,
		});
		expect(zip).not.toEqual(aliased);
	});

	it("rewrites one compressed size", async () => {
		expect(
			centralRecord(withCompressedSize(await twoEntries(), "a", 7), "a").compressedSize,
		).toBe(7);
	});

	it("rewrites one local header offset", async () => {
		expect(
			centralRecord(withLocalHeaderOffset(await twoEntries(), "a", 1), "a").localHeaderOffset,
		).toBe(1);
	});

	it("reverses the directory without moving it", async () => {
		const zip = await twoEntries();
		const reversed = readCentralDirectory(reverseCentralDirectory(zip));
		expect(reversed.records.map((r) => r.name)).toEqual(["b", "a"]);
		expect(reversed.offset).toBe(readCentralDirectory(zip).offset);
	});
});
