import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openZip } from "../../zip";
import { writeZip } from "../zip";

// F3.1 acceptance: whatever writeZip packs must re-read byte-identically through the reader
// (openZip, F1.3). These tests are the round-trip proof — for both the stored and the deflated
// path — plus the determinism and the ZIP64/duplicate guards.

const enc = new TextEncoder();
const bytes = (s: string): Uint8Array => enc.encode(s);

/** Read every entry back out of a written archive as a name -> bytes map. */
async function readBack(archive: Uint8Array): Promise<Map<string, Uint8Array>> {
	const zip = openZip(archive);
	const out = new Map<string, Uint8Array>();
	for (const name of zip.entries.keys()) out.set(name, await zip.read(name));
	return out;
}

describe("writeZip — round-trips through openZip", () => {
	it("stores tiny/incompressible parts (method 0) and reads them back verbatim", async () => {
		const data = bytes("hi"); // 2 bytes: deflating can only make it bigger
		const archive = await writeZip([{ name: "a.txt", data }]);
		const zip = openZip(archive);
		expect(zip.entries.get("a.txt")?.method).toBe(0);
		expect(Array.from(await zip.read("a.txt"))).toEqual(Array.from(data));
	});

	it("deflates compressible parts (method 8) and inflates them back identically", async () => {
		const data = bytes("A".repeat(2000)); // highly compressible → deflate is a clear win
		const archive = await writeZip([{ name: "big.txt", data }]);
		const zip = openZip(archive);
		const entry = zip.entries.get("big.txt");
		expect(entry?.method).toBe(8);
		expect(entry?.compressedSize).toBeLessThan(data.length);
		expect(entry?.uncompressedSize).toBe(data.length);
		expect(Array.from(await zip.read("big.txt"))).toEqual(Array.from(data));
	});

	it("round-trips many parts, preserving names and bytes", async () => {
		const parts = [
			{ name: "[Content_Types].xml", data: bytes("<Types/>") },
			{ name: "_rels/.rels", data: bytes("<Relationships/>") },
			{ name: "xl/workbook.xml", data: bytes("X".repeat(500)) },
			{ name: "xl/worksheets/sheet1.xml", data: bytes("<worksheet/>") },
		];
		const back = await readBack(await writeZip(parts));
		expect([...back.keys()]).toEqual(parts.map((p) => p.name));
		for (const part of parts) {
			expect(Array.from(back.get(part.name) ?? [])).toEqual(Array.from(part.data));
		}
	});

	it("round-trips an empty part (stored, zero-length)", async () => {
		const archive = await writeZip([{ name: "empty", data: new Uint8Array(0) }]);
		const zip = openZip(archive);
		expect(zip.entries.get("empty")?.method).toBe(0);
		expect((await zip.read("empty")).length).toBe(0);
	});

	it("round-trips non-ASCII (UTF-8) payload bytes", async () => {
		const data = bytes("héllo — wörld 😀 café");
		const back = await readBack(await writeZip([{ name: "u.txt", data }]));
		expect(Array.from(back.get("u.txt") ?? [])).toEqual(Array.from(data));
	});
});

describe("writeZip — determinism", () => {
	it("produces byte-identical output for identical input", async () => {
		const parts = [
			{ name: "a", data: bytes("A".repeat(2000)) },
			{ name: "b", data: bytes("hi") },
		];
		const first = await writeZip(parts);
		const second = await writeZip(parts);
		expect(Array.from(first)).toEqual(Array.from(second));
	});
});

describe("writeZip — guards", () => {
	it("rejects duplicate entry names (the reader would refuse the package)", async () => {
		const parts = [
			{ name: "dup", data: bytes("one") },
			{ name: "dup", data: bytes("two") },
		];
		await expect(writeZip(parts)).rejects.toThrow(/duplicate zip entry name/);
	});

	it("refuses an entry count that would require ZIP64", async () => {
		// 0xffff entries — the length guard fires before any compression, so this is cheap.
		const many = new Array(0xffff).fill({ name: "x", data: new Uint8Array(0) });
		const err = await writeZip(many).then(
			() => undefined,
			(e) => e,
		);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("unsupported");
	});

	it("refuses a name longer than the u16 name-length field", async () => {
		// > 65535 UTF-8 bytes: u16() would truncate the recorded length while the full name bytes
		// are still appended, so the reader would mislocate the payload. Reject instead.
		const err = await writeZip([{ name: "a".repeat(70000), data: bytes("x") }]).then(
			() => undefined,
			(e) => e,
		);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("unsupported");
	});

	it('rejects a name ending in "/" (the reader drops directory placeholders)', async () => {
		await expect(writeZip([{ name: "xl/", data: bytes("hello") }])).rejects.toThrow(
			/directory placeholder/,
		);
	});
});
