import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError, type XlsxErrorCode } from "../../errors";
import { openZip, storedSlices } from "../central-directory";
import { MAX_CLASSIC_ZIP_ENTRIES } from "../constants";

// The known parts in basic.xlsx; the same list also drives buffered-versus-streamed comparisons.
const PARTS = [
	"[Content_Types].xml",
	"_rels/.rels",
	"xl/_rels/workbook.xml.rels",
	"xl/sharedStrings.xml",
	"xl/styles.xml",
	"xl/workbook.xml",
	"xl/worksheets/sheet1.xml",
];

// Decode ZIP part bytes so XML content can be asserted as readable text.
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe("openZip — real fixture (stored entries)", () => {
	it("lists every part of basic.xlsx", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"));
		expect([...zip.entries.keys()].sort()).toEqual([...PARTS].sort());
	});

	it("reads parts back to their exact bytes", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"));

		const contentTypes = decode(await zip.read("[Content_Types].xml"));
		expect(contentTypes).toContain("<Types");
		expect(contentTypes).toContain("spreadsheetml.sheet.main+xml");

		const sheet = decode(await zip.read("xl/worksheets/sheet1.xml"));
		expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>');
		expect(sheet).toContain("<f>B1*2</f><v>84</v>");

		const shared = decode(await zip.read("xl/sharedStrings.xml"));
		expect(shared).toContain("<t>hello</t>");
		expect(shared).toContain("<t>world</t>");
	});

	it("reports membership and rejects unknown entries", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"));
		expect(zip.has("xl/workbook.xml")).toBe(true);
		expect(zip.has("nope.xml")).toBe(false);
		await expect(zip.read("nope.xml")).rejects.toThrow();
	});

	it("throws on a buffer that is not a zip", () => {
		expect(() => openZip(new Uint8Array([1, 2, 3, 4]))).toThrow(/not a zip/);
	});
});

// --- a configurable single-entry zip builder for the deflate / edge-case tests ---

// Keep test ZIP checksums independent from the production CRC utility so the tests cannot
// accidentally validate an implementation against itself.
const CRC_TABLE = (() => {
	const table = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		table[n] = c >>> 0;
	}
	return table;
})();

// Compute the independent CRC-32 written into synthetic ZIP headers and used in expectations.
function crc32(bytes: Uint8Array): number {
	let crc = 0xffffffff;
	for (let i = 0; i < bytes.length; i++) {
		crc = ((CRC_TABLE[(crc ^ (bytes[i] as number)) & 0xff] as number) ^ (crc >>> 8)) >>> 0;
	}
	return (crc ^ 0xffffffff) >>> 0;
}

// Encode the little-endian integers used by ZIP headers and directory records.
const u16 = (n: number): Uint8Array => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff]);
const u32 = (n: number): Uint8Array =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]);

// Join binary record fragments into headers, descriptors, directories, or a complete archive.
function concat(parts: Uint8Array[]): Uint8Array {
	const total = parts.reduce((sum, p) => sum + p.length, 0);
	const out = new Uint8Array(total);
	let offset = 0;
	for (const part of parts) {
		out.set(part, offset);
		offset += part.length;
	}
	return out;
}

// Produce a genuine raw-DEFLATE payload for ZIP compression method 8 without a test dependency.
async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	const stream = new Blob([data as unknown as BlobPart])
		.stream()
		.pipeThrough(new CompressionStream("deflate-raw"));
	return new Uint8Array(await new Response(stream).arrayBuffer());
}

interface ZipOptions {
	name: string;
	content: string;
	method?: number;
	/** general-purpose bit 3: zero the local sizes and append a trailing data descriptor */
	dataDescriptor?: boolean;
	descriptorSignature?: boolean;
	/** bytes of (zeroed) extra field in the local header */
	localExtra?: number;
	/** bytes of (zeroed) extra field in the central header */
	centralExtra?: number;
	/** bytes of (zeroed) comment in the central header */
	centralComment?: number;
	centralCompressedSize?: number;
	centralUncompressedSize?: number;
	centralCrc32?: number;
	centralLocalOffset?: number;
	/** replace the encoded payload while retaining metadata derived from `content` */
	payloadOverride?: Uint8Array;
	/** replace the payload with non-deflate garbage */
	corruptPayload?: boolean;
}

// Build a minimal single-entry ZIP whose individual fields can be overridden for hostile cases.
async function buildZip(o: ZipOptions): Promise<Uint8Array> {
	const enc = new TextEncoder();
	const raw = enc.encode(o.content);
	const method = o.method ?? 8;
	let payload = o.payloadOverride ?? (method === 8 ? await deflateRaw(raw) : raw);
	if (o.corruptPayload) payload = Uint8Array.from([0xff, 0xff, 0xff, 0xff]);
	const nameBytes = enc.encode(o.name);
	const crc = crc32(raw);
	const localExtra = new Uint8Array(o.localExtra ?? 0);
	const centralExtra = new Uint8Array(o.centralExtra ?? 0);
	const centralComment = new Uint8Array(o.centralComment ?? 0);
	const dd = o.dataDescriptor === true;
	const flags = dd ? 0x08 : 0;

	const local = concat([
		u32(0x04034b50),
		u16(20),
		u16(flags),
		u16(method),
		u16(0),
		u16(0x21),
		u32(dd ? 0 : crc),
		u32(dd ? 0 : payload.length),
		u32(dd ? 0 : raw.length),
		u16(nameBytes.length),
		u16(localExtra.length),
		nameBytes,
		localExtra,
	]);
	const blockParts = [local, payload];
	if (dd) {
		blockParts.push(
			o.descriptorSignature
				? concat([u32(0x08074b50), u32(crc), u32(payload.length), u32(raw.length)])
				: concat([u32(crc), u32(payload.length), u32(raw.length)]),
		);
	}
	const localBlock = concat(blockParts);

	const central = concat([
		u32(0x02014b50),
		u16(20),
		u16(20),
		u16(flags),
		u16(method),
		u16(0),
		u16(0x21),
		u32(o.centralCrc32 ?? crc),
		u32(o.centralCompressedSize ?? payload.length),
		u32(o.centralUncompressedSize ?? raw.length),
		u16(nameBytes.length),
		u16(centralExtra.length),
		u16(centralComment.length),
		u16(0),
		u16(0),
		u32(0),
		u32(o.centralLocalOffset ?? 0),
		nameBytes,
		centralExtra,
		centralComment,
	]);
	const eocd = concat([
		u32(0x06054b50),
		u16(0),
		u16(0),
		u16(1),
		u16(1),
		u32(central.length),
		u32(localBlock.length),
		u16(0),
	]);
	return concat([localBlock, central, eocd]);
}

// Bytes in the fixed EOCD record; an optional ZIP comment follows these 22 bytes.
const EOCD_FIXED_SIZE = 22;

// Clone an archive and alter one 16-bit EOCD field for disk, count, and sentinel tests.
function patchEocdU16(bytes: Uint8Array, fieldOffset: number, value: number): Uint8Array {
	const copy = bytes.slice();
	const eocd = copy.length - EOCD_FIXED_SIZE;
	new DataView(copy.buffer, copy.byteOffset, copy.byteLength).setUint16(
		eocd + fieldOffset,
		value,
		true,
	);
	return copy;
}

// Clone an archive and alter one 32-bit EOCD field for size, offset, and sentinel tests.
function patchEocdU32(bytes: Uint8Array, fieldOffset: number, value: number): Uint8Array {
	const copy = bytes.slice();
	const eocd = copy.length - EOCD_FIXED_SIZE;
	new DataView(copy.buffer, copy.byteOffset, copy.byteLength).setUint32(
		eocd + fieldOffset,
		value,
		true,
	);
	return copy;
}

// Read a valid 32-bit EOCD value so a corruption can be derived without hard-coded values.
function readEocdU32(bytes: Uint8Array, fieldOffset: number): number {
	const eocd = bytes.length - EOCD_FIXED_SIZE;
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(
		eocd + fieldOffset,
		true,
	);
}

// Mark the central entry as belonging to another disk to exercise multi-disk rejection.
function patchCentralDiskStart(bytes: Uint8Array, value: number): Uint8Array {
	const copy = bytes.slice();
	const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
	const eocd = copy.length - EOCD_FIXED_SIZE;
	const centralOffset = view.getUint32(eocd + 16, true);
	view.setUint16(centralOffset + 34, value, true);
	return copy;
}

// Insert a declared central-directory byte that is not a valid entry record.
function addDeclaredCentralTrailingByte(bytes: Uint8Array): Uint8Array {
	const oldEocd = bytes.length - EOCD_FIXED_SIZE;
	const oldSize = readEocdU32(bytes, 12);
	const copy = new Uint8Array(bytes.length + 1);
	copy.set(bytes.subarray(0, oldEocd));
	copy[oldEocd] = 0;
	copy.set(bytes.subarray(oldEocd), oldEocd + 1);
	const newEocd = oldEocd + 1;
	new DataView(copy.buffer).setUint32(newEocd + 12, oldSize + 1, true);
	return copy;
}

// Forge internally consistent counts and size that hide all but the first physical entry.
function declareOnlyFirstCentralEntry(bytes: Uint8Array): Uint8Array {
	const copy = bytes.slice();
	const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
	const eocd = copy.length - EOCD_FIXED_SIZE;
	const centralOffset = view.getUint32(eocd + 16, true);
	const nameLength = view.getUint16(centralOffset + 28, true);
	const extraLength = view.getUint16(centralOffset + 30, true);
	const commentLength = view.getUint16(centralOffset + 32, true);
	const firstRecordSize = 46 + nameLength + extraLength + commentLength;
	view.setUint16(eocd + 8, 1, true);
	view.setUint16(eocd + 10, 1, true);
	view.setUint32(eocd + 12, firstRecordSize, true);
	return copy;
}

// Add an EOCD-shaped ZIP comment to ensure the backward scan does not accept the fake record.
function appendFakeEmptyEocdAsComment(bytes: Uint8Array): Uint8Array {
	const oldEocd = bytes.length - EOCD_FIXED_SIZE;
	const fakeEocd = concat([
		u32(0x06054b50),
		u16(0),
		u16(0),
		u16(0),
		u16(0),
		u32(0),
		u32(0),
		u16(0),
	]);
	const copy = concat([bytes, fakeEocd]);
	new DataView(copy.buffer).setUint16(oldEocd + 20, fakeEocd.length, true);
	return copy;
}

// Assert that archive opening fails through the public typed-error contract, with no manual catch.
async function expectOpenZipError(
	bytes: Uint8Array,
	code: XlsxErrorCode,
	message?: RegExp,
): Promise<void> {
	const opening = Promise.resolve().then(() => openZip(bytes));
	await expect(opening).rejects.toBeInstanceOf(XlsxError);
	await expect(opening).rejects.toMatchObject({
		code,
		...(message === undefined ? {} : { message: expect.stringMatching(message) }),
	});
}

// Drain and join streamed chunks; reaching natural EOF also triggers final size and CRC checks.
async function collectChunks(chunks: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
	const parts: Uint8Array[] = [];
	let total = 0;
	for await (const chunk of chunks) {
		parts.push(chunk);
		total += chunk.byteLength;
	}
	const out = new Uint8Array(total);
	let offset = 0;
	for (const part of parts) {
		out.set(part, offset);
		offset += part.byteLength;
	}
	return out;
}

// Reuse the real openZip option type so test helpers stay aligned with the production API.
type OpenZipOptions = Parameters<typeof openZip>[1];

// Require buffered and fully consumed streaming reads to reject the same malformed entry.
async function expectReadAndStreamError(
	bytes: Uint8Array,
	name: string,
	options: OpenZipOptions,
	code: XlsxErrorCode,
): Promise<void> {
	await expect(openZip(bytes, options).read(name)).rejects.toMatchObject({ code });
	await expect(collectChunks(openZip(bytes, options).readStream(name))).rejects.toMatchObject({
		code,
	});
}

// Build a deterministic uncompressed DEFLATE block whose bit-flipped data remains decodable.
function rawStoredDeflate(data: Uint8Array): Uint8Array {
	if (data.byteLength > 0xffff) throw new Error("test DEFLATE block is too large");
	return concat([
		Uint8Array.of(0x01),
		u16(data.byteLength),
		u16(~data.byteLength & 0xffff),
		data,
	]);
}

// Flip only a payload bit, leaving ZIP sizes and headers intact so CRC is the isolated detector.
function flipEntryPayloadByte(bytes: Uint8Array, payloadOffset: number): Uint8Array {
	const copy = bytes.slice();
	const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
	const nameLength = view.getUint16(26, true);
	const extraLength = view.getUint16(28, true);
	const byteOffset = 30 + nameLength + extraLength + payloadOffset;
	const value = copy[byteOffset];
	if (value === undefined) throw new Error("test payload offset is out of bounds");
	copy[byteOffset] = value ^ 1;
	return copy;
}

// Write the test's safe integer values as little-endian ZIP64 fields with a zero high word.
function setU64(view: DataView, offset: number, value: number): void {
	view.setUint32(offset, value, true);
	view.setUint32(offset + 4, 0, true);
}

// A structurally genuine ZIP64 archive with more entries than classic ZIP can represent. Building
// directly into one buffer keeps this regression fast and avoids tens of thousands of allocations.
function buildOversizedZip64(): Uint8Array {
	const entryCount = MAX_CLASSIC_ZIP_ENTRIES + 2;
	const nameLength = 5;
	const localRecordSize = 30 + nameLength;
	const centralRecordSize = 46 + nameLength;
	const cdOffset = entryCount * localRecordSize;
	const cdSize = entryCount * centralRecordSize;
	const zip64EocdSize = 56;
	const zip64LocatorSize = 20;
	const zip64EocdOffset = cdOffset + cdSize;
	const locatorOffset = zip64EocdOffset + zip64EocdSize;
	const eocdOffset = locatorOffset + zip64LocatorSize;
	const bytes = new Uint8Array(eocdOffset + EOCD_FIXED_SIZE);
	const view = new DataView(bytes.buffer);

	for (let i = 0; i < entryCount; i++) {
		const localOffset = i * localRecordSize;
		view.setUint32(localOffset, 0x04034b50, true);
		view.setUint16(localOffset + 4, 20, true);
		view.setUint16(localOffset + 12, 0x21, true);
		view.setUint16(localOffset + 26, nameLength, true);

		const name = `e${i.toString(36).padStart(4, "0")}`;
		for (let j = 0; j < nameLength; j++) {
			bytes[localOffset + 30 + j] = name.charCodeAt(j);
		}

		const centralOffset = cdOffset + i * centralRecordSize;
		view.setUint32(centralOffset, 0x02014b50, true);
		view.setUint16(centralOffset + 4, 20, true);
		view.setUint16(centralOffset + 6, 20, true);
		view.setUint16(centralOffset + 14, 0x21, true);
		view.setUint16(centralOffset + 28, nameLength, true);
		view.setUint32(centralOffset + 42, localOffset, true);
		bytes.set(bytes.subarray(localOffset + 30, localOffset + 35), centralOffset + 46);
	}

	view.setUint32(zip64EocdOffset, 0x06064b50, true);
	setU64(view, zip64EocdOffset + 4, 44);
	view.setUint16(zip64EocdOffset + 12, 45, true);
	view.setUint16(zip64EocdOffset + 14, 45, true);
	setU64(view, zip64EocdOffset + 24, entryCount);
	setU64(view, zip64EocdOffset + 32, entryCount);
	setU64(view, zip64EocdOffset + 40, cdSize);
	setU64(view, zip64EocdOffset + 48, cdOffset);

	view.setUint32(locatorOffset, 0x07064b50, true);
	setU64(view, locatorOffset + 8, zip64EocdOffset);
	view.setUint32(locatorOffset + 16, 1, true);

	view.setUint32(eocdOffset, 0x06054b50, true);
	view.setUint16(eocdOffset + 8, 0xffff, true);
	view.setUint16(eocdOffset + 10, 0xffff, true);
	view.setUint32(eocdOffset + 12, cdSize, true);
	view.setUint32(eocdOffset + 16, cdOffset, true);
	return bytes;
}

describe("openZip — deflate and real-world layouts", () => {
	it("inflates a plain deflate (method 8) entry", async () => {
		const content = "deflate me ".repeat(500);
		const zip = openZip(await buildZip({ name: "big.xml", content }));
		expect(zip.entries.get("big.xml")?.method).toBe(8);
		expect(decode(await zip.read("big.xml"))).toBe(content);
	});

	it("reads a data-descriptor entry (GP bit 3, zeroed local sizes, real central sizes)", async () => {
		const content = "rows ".repeat(400);
		for (const descriptorSignature of [false, true]) {
			const zip = openZip(
				await buildZip({
					name: "sheet.xml",
					content,
					dataDescriptor: true,
					descriptorSignature,
				}),
			);
			expect(decode(await zip.read("sheet.xml"))).toBe(content);
		}
	});

	it("uses the local (not central) extra-field length to locate the data", async () => {
		const content = "extra fields differ ".repeat(50);
		const zip = openZip(
			await buildZip({ name: "p.xml", content, localExtra: 12, centralExtra: 0 }),
		);
		expect(decode(await zip.read("p.xml"))).toBe(content);
	});

	it("caps inflate at the declared uncompressed size, typed part-too-large (F9.7)", async () => {
		// A stream that produces more than its declared uncompressedSize is a size/bomb rejection
		// (part-too-large), NOT a corrupt stream — the inflate aborts while streaming.
		const zip = openZip(
			await buildZip({
				name: "bomb.xml",
				content: "x".repeat(10000),
				centralUncompressedSize: 16,
			}),
		);
		await expect(zip.read("bomb.xml")).rejects.toMatchObject({ code: "part-too-large" });
	});

	it("wraps inflate failures with the part name", async () => {
		const zip = openZip(
			await buildZip({ name: "bad.xml", content: "whatever", corruptPayload: true }),
		);
		await expect(zip.read("bad.xml")).rejects.toThrow(
			/corrupt zip: failed to inflate bad\.xml/,
		);
		await expect(zip.read("bad.xml")).rejects.toMatchObject({ code: "corrupt-zip" });
		// The streaming inflater owns the same runtime-error boundary as the buffered inflater.
		await expect(collectChunks(zip.readStream("bad.xml"))).rejects.toThrow(
			/corrupt zip: failed to inflate bad\.xml/,
		);
		await expect(collectChunks(zip.readStream("bad.xml"))).rejects.toMatchObject({
			code: "corrupt-zip",
		});
	});

	it("returns empty bytes for a zero-length deflate entry", async () => {
		const zip = openZip(
			await buildZip({
				name: "empty.xml",
				content: "",
				centralCompressedSize: 0,
				centralUncompressedSize: 0,
			}),
		);
		expect((await zip.read("empty.xml")).length).toBe(0);
	});

	it("rejects ZIP64 archives with an explicit error", async () => {
		const zip = await buildZip({ name: "p.xml", content: "x", centralLocalOffset: 0xffffffff });
		await expectOpenZipError(zip, "unsupported", /ZIP64/);
	});

	it("rejects a part whose declared size exceeds maxPartBytes", async () => {
		const bytes = await buildZip({
			name: "big.xml",
			content: "x",
			centralUncompressedSize: 10000,
		});
		const zip = openZip(bytes, { maxPartBytes: 100 });
		await expect(zip.read("big.xml")).rejects.toMatchObject({ code: "part-too-large" });
	});
});

describe("openZip — central-directory validation (F11.2)", () => {
	// ZIP64 uses all-ones sentinel values in classic EOCD fields. These are recognized extension
	// markers, so they must fail as `unsupported` before any entry walk or count comparison.
	it.each([
		["entries on this disk", 8],
		["total entries", 10],
	])("rejects a ZIP64 sentinel in the EOCD %s field", async (_field, offset) => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(patchEocdU16(valid, offset, 0xffff), "unsupported", /ZIP64/);
	});

	it.each([
		["central-directory size", 12],
		["central-directory offset", 16],
	])("rejects a ZIP64 sentinel in the EOCD %s field", async (_field, offset) => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(patchEocdU32(valid, offset, 0xffffffff), "unsupported", /ZIP64/);
	});

	// Any disk metadata means the package is split across files. Returning its local subset would
	// look like a valid but incomplete workbook, so every split marker is rejected at open time.
	it.each([
		["current disk", 4],
		["central-directory start disk", 6],
	])("rejects a nonzero EOCD %s field", async (_field, offset) => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(patchEocdU16(valid, offset, 1), "unsupported", /multi-disk/);
	});

	it("rejects unequal per-disk and total EOCD entry counts", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(patchEocdU16(valid, 8, 0), "unsupported", /multi-disk/);
	});

	it("rejects a central-directory entry stored on another disk", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(patchCentralDiskStart(valid, 1), "unsupported", /multi-disk/);
	});

	// The EOCD size/offset pair defines the only bytes the entry walker may consume. These mutations
	// cover a short span, an EOCD overlap, undeclared trailing data, and a shifted starting boundary.
	it("rejects an under-declared central-directory span", async () => {
		const valid = await buildZip({
			name: "p.xml",
			content: "x",
			method: 0,
			centralExtra: 2,
			centralComment: 2,
		});
		const realSize = readEocdU32(valid, 12);
		await expectOpenZipError(patchEocdU32(valid, 12, realSize - 1), "corrupt-zip");
	});

	it("rejects a declared central-directory span that overlaps the EOCD", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		const realSize = readEocdU32(valid, 12);
		await expectOpenZipError(patchEocdU32(valid, 12, realSize + 1), "corrupt-zip");
	});

	it("rejects trailing bytes inside the declared central-directory span", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(addDeclaredCentralTrailingByte(valid), "corrupt-zip");
	});

	it("rejects a shifted central-directory offset even when the declared end is unchanged", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		const realOffset = readEocdU32(valid, 16);
		const realSize = readEocdU32(valid, 12);
		const shifted = patchEocdU32(patchEocdU32(valid, 16, realOffset - 1), 12, realSize + 1);
		await expectOpenZipError(shifted, "corrupt-zip");
	});

	// Entry counts must agree with the physical directory. The second mutation shrinks count and
	// size together; it must not return a believable map that silently omits later entries.
	it("rejects a zero entry count that would silently discard a declared directory", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		const zeroCounts = patchEocdU16(patchEocdU16(valid, 8, 0), 10, 0);
		await expectOpenZipError(zeroCounts, "corrupt-zip");
	});

	it("rejects consistent forged metadata that hides later central-directory entries", async () => {
		const valid = await loadFixture("basic.xlsx");
		expect(openZip(valid).entries.size).toBeGreaterThan(1);
		await expectOpenZipError(declareOnlyFirstCentralEntry(valid), "corrupt-zip");
	});

	// Both the real and fake EOCD signatures land at valid comment boundaries. Choosing the fake one
	// must fail closed because its empty directory does not immediately precede that record.
	it("fails closed when an EOCD-shaped comment suffix would hide every entry", async () => {
		const valid = await buildZip({ name: "p.xml", content: "x", method: 0 });
		await expectOpenZipError(appendFakeEmptyEocdAsComment(valid), "corrupt-zip");
	});

	it("accepts a valid empty classic ZIP", () => {
		const emptyZip = concat([
			u32(0x06054b50),
			u16(0),
			u16(0),
			u16(0),
			u16(0),
			u32(0),
			u32(0),
			u16(0),
		]);
		expect(openZip(emptyZip).entries.size).toBe(0);
	});

	// This is a real 65,536-entry ZIP64 layout, not only a forged EOCD field. It proves the reader
	// rejects the writer's shared classic-ZIP boundary before silently dropping the final entry.
	it("rejects a genuine archive above the shared classic-ZIP entry ceiling", async () => {
		const zip64 = buildOversizedZip64();
		const eocd = zip64.length - EOCD_FIXED_SIZE;
		const view = new DataView(zip64.buffer);
		const zip64Eocd = eocd - 20 - 56;
		expect(view.getUint32(zip64Eocd, true)).toBe(0x06064b50);
		expect(view.getUint32(zip64Eocd + 32, true)).toBe(MAX_CLASSIC_ZIP_ENTRIES + 2);
		expect(view.getUint32(zip64Eocd + 36, true)).toBe(0);
		expect(view.getUint32(eocd - 20, true)).toBe(0x07064b50);
		expect(view.getUint16(eocd + 10, true)).toBe(0xffff);
		await expectOpenZipError(zip64, "unsupported", /ZIP64/);
	});
});

describe("openZip — completed-entry integrity (F11.3)", () => {
	it("reads the central CRC and verifies valid stored, deflated, and data-descriptor entries", async () => {
		const enc = new TextEncoder();
		for (const options of [
			{ name: "stored.txt", content: "stored checksum", method: 0 },
			{ name: "deflated.txt", content: "deflated checksum", method: 8 },
			{ name: "descriptor.txt", content: "descriptor checksum", dataDescriptor: true },
		]) {
			const bytes = await buildZip(options);
			const zip = openZip(bytes, { verifyChecksums: true });
			expect(zip.entries.get(options.name)?.crc32).toBe(crc32(enc.encode(options.content)));
			expect(decode(await zip.read(options.name))).toBe(options.content);
			expect(decode(await collectChunks(zip.readStream(options.name)))).toBe(options.content);
		}
	});

	it.each([
		{ kind: "stored", method: 0, flipOffset: 0 },
		{ kind: "deflated", method: 8, flipOffset: 5 },
	])(
		"detects a bit-flipped $kind payload only when CRC verification is enabled",
		async ({ kind, method, flipOffset }) => {
			const enc = new TextEncoder();
			const name = `${kind}.txt`;
			const content = `${kind} payload`;
			const raw = enc.encode(content);
			const bytes = await buildZip({
				name,
				content,
				method,
				...(method === 8 ? { payloadOverride: rawStoredDeflate(raw) } : {}),
			});
			const damaged = flipEntryPayloadByte(bytes, flipOffset);

			// CRC verification is deliberately opt-in for 1.0.1: the same-size corruption remains
			// readable by default, preserving the previous behavior.
			expect(await openZip(damaged).read(name)).not.toEqual(raw);
			expect(await collectChunks(openZip(damaged).readStream(name))).not.toEqual(raw);

			await expectReadAndStreamError(damaged, name, { verifyChecksums: true }, "corrupt-zip");
		},
	);

	it("does not treat an expected CRC of zero as verification being disabled", async () => {
		const name = "zero-crc.txt";
		const bytes = await buildZip({ name, content: "nonempty", method: 0, centralCrc32: 0 });
		expect(decode(await openZip(bytes).read(name))).toBe("nonempty");
		await expectReadAndStreamError(bytes, name, { verifyChecksums: true }, "corrupt-zip");
	});

	it("rejects deflate output that ends before its declared uncompressed size", async () => {
		const name = "short.xml";
		const content = "short output";
		const bytes = await buildZip({
			name,
			content,
			centralUncompressedSize: content.length + 1,
		});
		await expectReadAndStreamError(bytes, name, undefined, "corrupt-zip");
		await expectReadAndStreamError(bytes, name, { verifyChecksums: true }, "corrupt-zip");
	});

	it("rejects a stored entry whose compressed and uncompressed sizes disagree", async () => {
		const name = "stored.bin";
		const bytes = await buildZip({
			name,
			content: "stored bytes",
			method: 0,
			centralUncompressedSize: 13,
		});
		await expectReadAndStreamError(bytes, name, undefined, "corrupt-zip");
		await expectReadAndStreamError(bytes, name, { verifyChecksums: true }, "corrupt-zip");
	});

	it("checks the zero-compressed deflate shortcut against the declared output size", async () => {
		const name = "empty.xml";
		const bytes = await buildZip({
			name,
			content: "",
			centralCompressedSize: 0,
			centralUncompressedSize: 1,
		});
		await expectReadAndStreamError(bytes, name, undefined, "corrupt-zip");
	});

	it("verifies genuine empty stored and deflated entries", async () => {
		for (const method of [0, 8]) {
			const name = `empty-${method}.xml`;
			const bytes = await buildZip({ name, content: "", method });
			const zip = openZip(bytes, { verifyChecksums: true });
			expect(await zip.read(name)).toEqual(new Uint8Array(0));
			expect(await collectChunks(zip.readStream(name))).toEqual(new Uint8Array(0));
		}
	});

	it.each([
		{ kind: "stored", method: 0, flipOffset: 0 },
		{ kind: "deflated", method: 8, flipOffset: 5 },
	])(
		"makes no checksum claim when a $kind stream is cancelled before EOF",
		async ({ kind, method, flipOffset }) => {
			const name = `cancel-${kind}.txt`;
			const content = `cancel ${kind}`;
			const raw = new TextEncoder().encode(content);
			const valid = await buildZip({
				name,
				content,
				method,
				...(method === 8 ? { payloadOverride: rawStoredDeflate(raw) } : {}),
			});
			const damaged = flipEntryPayloadByte(valid, flipOffset);
			expect(() => openZip(damaged, { verifyChecksums: true })).not.toThrow();

			const stream = openZip(damaged, { verifyChecksums: true }).readStream(name);
			const first = await stream.next();
			expect(first.done).toBe(false);
			await expect(stream.return(undefined)).resolves.toMatchObject({ done: true });
		},
	);
});

// F9.7 — the zip-bomb guards are ON BY DEFAULT (no opt-in): an absolute per-part ceiling plus a
// compression-ratio cap over an 8 MiB floor. Both are caller-tunable / disable-able.
describe("openZip — decompression-bomb defaults (F9.7)", () => {
	it("catches a highly-compressible bomb by DEFAULT (no options), typed", async () => {
		// ~9 MiB of one repeated char deflates to a few hundred bytes — a >8 MiB output at a
		// ~10000:1 ratio. The default guard (8 MiB floor / 300× ratio) must abort it typed.
		const zip = openZip(
			await buildZip({ name: "bomb.xml", content: "A".repeat(9 * 1024 * 1024) }),
		);
		await expect(zip.read("bomb.xml")).rejects.toMatchObject({ code: "part-too-large" });
	});

	it("lets a normal part through under the default guards", async () => {
		// A realistic ~50:1 part well under the floor reads fine with NO options passed.
		const content = "<row><c><v>123.45</v></c></row>".repeat(2000);
		const zip = openZip(await buildZip({ name: "sheet.xml", content }));
		expect(decode(await zip.read("sheet.xml"))).toBe(content);
	});

	it("can be disabled with Infinity (caller opts out of both guards)", async () => {
		const bomb = await buildZip({ name: "bomb.xml", content: "A".repeat(9 * 1024 * 1024) });
		const zip = openZip(bomb, {
			maxPartBytes: Number.POSITIVE_INFINITY,
			maxCompressionRatio: Number.POSITIVE_INFINITY,
		});
		expect((await zip.read("bomb.xml")).length).toBe(9 * 1024 * 1024);
	});

	it("enforces maxPartBytes against a STORED entry's real size, not its lying uncompressedSize", async () => {
		// The pre-F9.7 gap: a STORED (method 0) part returns compressedSize bytes, but the guard
		// only checked uncompressedSize — which an attacker sets to 0 to slip a large part through.
		const bytes = await buildZip({
			name: "stored.bin",
			content: "S".repeat(5000),
			method: 0,
			centralUncompressedSize: 0, // the lie
		});
		const guarded = openZip(bytes, { maxPartBytes: 1000 });
		await expect(guarded.read("stored.bin")).rejects.toMatchObject({ code: "part-too-large" });
		// Without the tight cap, F11.3 reaches EOF and rejects the lying size as corrupt.
		const open = openZip(bytes);
		await expect(open.read("stored.bin")).rejects.toMatchObject({ code: "corrupt-zip" });
	});
});

describe("openZip — entry policy", () => {
	it("rejects a zip with duplicate entry names", async () => {
		const bytes = await loadFixture("edge-duplicate-entry.xlsx");
		await expectOpenZipError(bytes, "corrupt-zip", /duplicate entry name/);
	});

	it("skips directory placeholder entries, keeping real parts", async () => {
		const zip = openZip(await loadFixture("edge-with-directory.xlsx"));
		expect(zip.entries.has("sub/")).toBe(false);
		expect(zip.has("keep.xml")).toBe(true);
	});
});

describe("readStream", () => {
	it("streams the same bytes as read() for every part", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"));
		for (const part of PARTS) {
			expect(decode(await collectChunks(zip.readStream(part)))).toBe(
				decode(await zip.read(part)),
			);
		}
	});

	// F11.98: a stored entry arrives in 16 KiB slices, within the 64 KiB the task allows, the size of
	// Node's inflate chunks.
	const SLICE = 16 * 1024;
	// `n` letters whose order shifts between slices, so a reordered or repeated slice cannot join back.
	const letters = (n: number): string => {
		let out = "";
		for (let i = 0; i < n; i++) out += String.fromCharCode(97 + ((i * 7919) % 26));
		return out;
	};
	// Drain a stream, keeping each chunk's size; natural EOF runs the size and CRC checks.
	async function chunkSizes(chunks: AsyncIterable<Uint8Array>, sizes: number[] = []) {
		for await (const chunk of chunks) sizes.push(chunk.byteLength);
		return sizes;
	}

	it.each([
		{ size: 0, slices: [] },
		{ size: 100, slices: [100] },
		{ size: SLICE, slices: [SLICE] },
		{ size: SLICE + 1, slices: [SLICE, 1] },
		{ size: 2 * SLICE, slices: [SLICE, SLICE] },
		{ size: 3 * SLICE + 123, slices: [SLICE, SLICE, SLICE, 123] },
	])(
		"streams a $size-byte stored entry as $slices.length slice(s) that join to it (F11.98)",
		async ({ size, slices }) => {
			const name = "stored.xml";
			const content = letters(size);
			const zip = openZip(await buildZip({ name, content, method: 0 }), {
				verifyChecksums: true,
			});
			expect(await chunkSizes(zip.readStream(name))).toEqual(slices);
			expect(decode(await collectChunks(zip.readStream(name)))).toBe(content);
			expect(decode(await zip.read(name))).toBe(content);
		},
	);

	it.each([
		{ slice: "first", offset: 5 },
		{ slice: "second", offset: SLICE + 7 },
		{ slice: "last", offset: 3 * SLICE + 5 },
	])(
		"checks a stored entry's CRC over every slice, after the last, with a flip in the $slice (F11.98)",
		async ({ offset }) => {
			const name = "stored.xml";
			const content = letters(3 * SLICE + 123);
			const damaged = flipEntryPayloadByte(
				await buildZip({ name, content, method: 0 }),
				offset,
			);
			// Off by default: the damaged bytes read, at full length.
			const raw = await collectChunks(openZip(damaged).readStream(name));
			expect(raw.byteLength).toBe(content.length);
			expect(decode(raw)).not.toBe(content);
			// On: every slice arrives, then natural EOF fails the CRC.
			const sizes: number[] = [];
			const stream = openZip(damaged, { verifyChecksums: true }).readStream(name);
			await expect(chunkSizes(stream, sizes)).rejects.toMatchObject({ code: "corrupt-zip" });
			expect(sizes).toEqual([SLICE, SLICE, SLICE, 123]);
			await expectReadAndStreamError(damaged, name, { verifyChecksums: true }, "corrupt-zip");
		},
	);

	it("checks a stored entry's declared size after its last slice (F11.98)", async () => {
		const name = "stored.xml";
		const content = letters(2 * SLICE + 9);
		for (const declared of [content.length - 1, content.length + 1]) {
			const bytes = await buildZip({
				name,
				content,
				method: 0,
				centralUncompressedSize: declared,
			});
			const sizes: number[] = [];
			await expect(chunkSizes(openZip(bytes).readStream(name), sizes)).rejects.toMatchObject({
				code: "corrupt-zip",
			});
			expect(sizes).toEqual([SLICE, SLICE, 9]);
			await expectReadAndStreamError(bytes, name, undefined, "corrupt-zip");
		}
	});
});

describe("storedSlices (F11.98)", () => {
	const SLICE = 16 * 1024;

	it("yields nothing for an empty part", () => {
		expect([...storedSlices(new Uint8Array(0))]).toEqual([]);
	});

	it("yields a part of at most 16 KiB as one view of it", () => {
		for (const size of [1, 100, SLICE]) {
			const payload = new Uint8Array(size + 10).subarray(3, 3 + size);
			const slices = [...storedSlices(payload)];
			expect(slices).toHaveLength(1);
			expect(slices[0]?.buffer).toBe(payload.buffer);
			expect([slices[0]?.byteOffset, slices[0]?.byteLength]).toEqual([3, size]);
		}
	});

	it("cuts a larger part into 16 KiB views that join back to it, copying nothing", () => {
		for (const [size, lengths] of [
			[SLICE + 1, [SLICE, 1]],
			[2 * SLICE, [SLICE, SLICE]],
			[3 * SLICE + 123, [SLICE, SLICE, SLICE, 123]],
		] as const) {
			const backing = Uint8Array.from({ length: size + 10 }, (_, i) => (i * 31) & 0xff);
			const payload = backing.subarray(3, 3 + size);
			const slices = [...storedSlices(payload)];
			expect(slices.map((s) => s.byteLength)).toEqual(lengths);
			expect(slices.map((s) => s.byteOffset)).toEqual(lengths.map((_, i) => 3 + i * SLICE));
			for (const slice of slices) expect(slice.buffer).toBe(backing.buffer);
			const joined = new Uint8Array(size);
			let at = 0;
			for (const slice of slices) {
				joined.set(slice, at);
				at += slice.byteLength;
			}
			expect(joined).toEqual(payload);
		}
	});
});
