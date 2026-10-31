import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { detectSpreadsheetFormat } from "../../reader/detect";
import { openXlsx, streamSheetRows } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import { writeZip } from "../../writer/zip";
import { type ClassicEnd, isRedundantZip64End, openZip, smallU64 } from "../central-directory";

// F11.35: an archive whose ZIP64 end record and locator only repeat the classic end record opens as
// classic, as in 1.0.0. Every other ZIP64 shape stays typed `unsupported`, each from its own site.

const EOCD_SIZE = 22;
const RECORD_SIZE = 56;
const LOCATOR_SIZE = 20;

// The detail of each ZIP64 rejection site, so no pin can pass through another site.
const AT_EOCD =
	"ZIP64 archives are not supported: the end-of-central-directory record holds a ZIP64 sentinel";
const AT_END_RECORD =
	"ZIP64 archives are not supported: the ZIP64 end record does not repeat the end-of-central-directory record";
const atCentral = (offset: number) =>
	`ZIP64 archives are not supported: the central directory entry at offset ${offset} holds a ZIP64 sentinel`;

const PARTS = [
	{ name: "a.xml", data: new TextEncoder().encode("<a>first part</a>") },
	{ name: "b.xml", data: new TextEncoder().encode("<b>second part</b>") },
];

function viewOf(bytes: Uint8Array): DataView {
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/** Changes to a built trailer: u32 words of the record or the locator, by offset within it. */
interface TrailerPatch {
	readonly record?: Readonly<Record<number, number>>;
	readonly locator?: Readonly<Record<number, number>>;
	/** Bytes of extensible data after the 56-byte record; its size field counts them. */
	readonly extensible?: number;
	/** Keep only this many record bytes, so the locator follows a short gap. */
	readonly keep?: number;
	/** Zero bytes between the record and the locator. */
	readonly junk?: number;
}

/**
 * A copy of the classic archive `bytes` (no comment) with a ZIP64 end record and locator before its
 * end record, as bsdtar's `zip:zip64` writes them; `patch` then changes the trailer.
 *
 * @example openZip(withZip64End(zip)).has("a.xml") // true: the record repeats the classic values
 * @example openZip(withZip64End(zip, { record: { 32: 3 } })) // throws XlsxError "unsupported"
 */
function withZip64End(bytes: Uint8Array, patch: TrailerPatch = {}): Uint8Array {
	const eocd = bytes.length - EOCD_SIZE;
	const src = viewOf(bytes);
	const entriesThisDisk = src.getUint16(eocd + 8, true);
	const entryCount = src.getUint16(eocd + 10, true);
	const cdSize = src.getUint32(eocd + 12, true);
	const cdOffset = src.getUint32(eocd + 16, true);
	const recordSize = RECORD_SIZE + (patch.extensible ?? 0);

	const record = new Uint8Array(recordSize);
	const rv = viewOf(record);
	rv.setUint32(0, 0x06064b50, true);
	rv.setUint32(4, recordSize - 12, true);
	rv.setUint16(12, 45, true);
	rv.setUint16(14, 45, true);
	rv.setUint32(24, entriesThisDisk, true);
	rv.setUint32(32, entryCount, true);
	rv.setUint32(40, cdSize, true);
	rv.setUint32(48, cdOffset, true);
	for (const [at, value] of Object.entries(patch.record ?? {})) {
		rv.setUint32(Number(at), value, true);
	}
	const kept = record.subarray(0, patch.keep ?? recordSize);

	const locator = new Uint8Array(LOCATOR_SIZE);
	const lv = viewOf(locator);
	lv.setUint32(0, 0x07064b50, true);
	lv.setUint32(8, eocd, true);
	lv.setUint32(16, 1, true);
	for (const [at, value] of Object.entries(patch.locator ?? {})) {
		lv.setUint32(Number(at), value, true);
	}

	const locatorAt = eocd + kept.length + (patch.junk ?? 0);
	const out = new Uint8Array(locatorAt + LOCATOR_SIZE + EOCD_SIZE);
	out.set(bytes.subarray(0, eocd));
	out.set(kept, eocd);
	out.set(locator, locatorAt);
	out.set(bytes.subarray(eocd), locatorAt + LOCATOR_SIZE);
	return out;
}

// Opening fails typed `unsupported` with exactly this detail, through the public error contract.
async function expectUnsupported(bytes: Uint8Array, message: string): Promise<void> {
	const opening = Promise.resolve().then(() => openZip(bytes));
	await expect(opening).rejects.toBeInstanceOf(XlsxError);
	await expect(opening).rejects.toMatchObject({ code: "unsupported", message });
}

/** Where the classic archive's fields lie, for patches relative to them. */
interface Layout {
	readonly count: number;
	readonly cdSize: number;
	readonly cdOffset: number;
	/** Where `withZip64End` puts the record: the classic end record's old offset. */
	readonly recordAt: number;
}

function layoutOf(classic: Uint8Array): Layout {
	const view = viewOf(classic);
	const eocd = classic.length - EOCD_SIZE;
	return {
		count: view.getUint16(eocd + 10, true),
		cdSize: view.getUint32(eocd + 12, true),
		cdOffset: view.getUint32(eocd + 16, true),
		recordAt: eocd,
	};
}

/**
 * A copy of the classic archive `bytes` (no comment) whose last central-directory record carries a
 * comment ending in `trailer`, so the trailer's last 20 bytes sit at `eocd - 20` inside the directory.
 *
 * @example withDirectoryComment(zip, locator) // a locator signature inside the central directory
 */
function withDirectoryComment(bytes: Uint8Array, trailer: Uint8Array): Uint8Array {
	const eocd = bytes.length - EOCD_SIZE;
	const src = viewOf(bytes);
	const count = src.getUint16(eocd + 10, true);
	let last = src.getUint32(eocd + 16, true);
	for (let i = 1; i < count; i++) {
		last +=
			46 +
			src.getUint16(last + 28, true) +
			src.getUint16(last + 30, true) +
			src.getUint16(last + 32, true);
	}
	const out = new Uint8Array(bytes.length + trailer.length);
	out.set(bytes.subarray(0, eocd));
	out.set(trailer, eocd);
	out.set(bytes.subarray(eocd), eocd + trailer.length);
	const view = viewOf(out);
	view.setUint16(last + 32, src.getUint16(last + 32, true) + trailer.length, true);
	view.setUint32(
		eocd + trailer.length + 12,
		src.getUint32(eocd + 12, true) + trailer.length,
		true,
	);
	return out;
}

async function readAll(bytes: Uint8Array, verifyChecksums: boolean): Promise<string[]> {
	const zip = openZip(bytes, { verifyChecksums });
	const out: string[] = [];
	for (const name of zip.entries.keys()) {
		out.push(`${name}=${new TextDecoder().decode(await zip.read(name))}`);
	}
	return out;
}

async function cellValues(bytes: Uint8Array): Promise<unknown[][]> {
	const workbook = await openXlsx(bytes);
	const rows = [];
	for await (const row of workbook.sheet(workbook.sheets[0]?.name ?? "").rows()) {
		rows.push(row.cells.map((cell) => cell.value));
	}
	return rows;
}

describe("openZip — a ZIP64 end record that repeats the classic one (F11.35)", () => {
	it("opens it as the classic archive, with checksums verified or not", async () => {
		const classic = await writeZip(PARTS);
		const bytes = withZip64End(classic);
		expect(bytes.length).toBe(classic.length + RECORD_SIZE + LOCATOR_SIZE);
		for (const verifyChecksums of [false, true]) {
			expect(await readAll(bytes, verifyChecksums)).toEqual(
				await readAll(classic, verifyChecksums),
			);
		}
	});

	it("reads bsdtar's zip:zip64 workbook verbatim, as minimal.xlsx reads", async () => {
		const bytes = await loadFixture("bsdtar-zip64-end.xlsx");
		const view = viewOf(bytes);
		const eocd = bytes.length - EOCD_SIZE;
		const cdEnd = view.getUint32(eocd + 16, true) + view.getUint32(eocd + 12, true);
		expect(eocd - cdEnd).toBe(RECORD_SIZE + LOCATOR_SIZE);
		expect(view.getUint32(cdEnd, true)).toBe(0x06064b50);
		expect(view.getUint32(eocd - LOCATOR_SIZE, true)).toBe(0x07064b50);

		for (const verifyChecksums of [false, true]) {
			expect([...openZip(bytes, { verifyChecksums }).entries.keys()]).toEqual([
				"[Content_Types].xml",
				"_rels/.rels",
				"xl/workbook.xml",
				"xl/_rels/workbook.xml.rels",
				"xl/worksheets/sheet1.xml",
			]);
			await readAll(bytes, verifyChecksums);
		}
		expect(await cellValues(bytes)).toEqual([[1, true]]);
		expect(await cellValues(bytes)).toEqual(
			await cellValues(await loadFixture("minimal.xlsx")),
		);
		const streamed = [];
		for await (const row of streamSheetRows(bytes, "Sheet1", { verifyChecksums: true })) {
			streamed.push(row.cells.map((cell) => cell.value));
		}
		expect(streamed).toEqual([[1, true]]);
		expect(await detectSpreadsheetFormat(bytes)).toBe("xlsx");
	});

	it("classifies a workbook with the redundant record as xlsx and reads it", async () => {
		const bytes = withZip64End(
			await writeXlsx({ sheets: [{ name: "Data", rows: [[7, "x"]] }] }),
		);
		expect(await detectSpreadsheetFormat(bytes)).toBe("xlsx");
		expect(await cellValues(bytes)).toEqual([[7, "x"]]);
	});

	it.each([
		["bsdtar's version fields", {}],
		["Info-ZIP's version made by (798)", { record: { 12: 798 | (45 << 16) } }],
		["a locator that counts no disks", { locator: { 16: 0 } }],
	])("accepts %s", async (_shape, patch: TrailerPatch) => {
		expect(openZip(withZip64End(await writeZip(PARTS), patch)).has("b.xml")).toBe(true);
	});
});

describe("openZip — ZIP64 end records that do not repeat the classic one (F11.35)", () => {
	// Each patch is relative to the classic archive's layout; the record's and the locator's fields
	// are named by their offsets (APPNOTE 4.3.14, 4.3.15).
	it.each<[string, (c: Layout) => TrailerPatch]>([
		["entries on this disk + 1", (c) => ({ record: { 24: c.count + 1 } })],
		["entries on this disk − 1", (c) => ({ record: { 24: c.count - 1 } })],
		["total entries + 1", (c) => ({ record: { 32: c.count + 1 } })],
		["total entries − 1", (c) => ({ record: { 32: c.count - 1 } })],
		["central-directory size − 1", (c) => ({ record: { 40: c.cdSize - 1 } })],
		["central-directory size + 1", (c) => ({ record: { 40: c.cdSize + 1 } })],
		["central-directory offset + 1", (c) => ({ record: { 48: c.cdOffset + 1 } })],
		["central-directory offset − 1", (c) => ({ record: { 48: c.cdOffset - 1 } })],
		["the record's disk 1", () => ({ record: { 16: 1 } })],
		["the record's central-directory disk 1", () => ({ record: { 20: 1 } })],
		["a bad record signature", () => ({ record: { 0: 0x06054b50 } })],
		["record size 60 in a 56-byte record", () => ({ record: { 4: 60 } })],
		["record size 60 with 16 bytes of extensible data", () => ({ extensible: 16 })],
		["an agreeing record, then 1 byte before the locator", () => ({ junk: 1 })],
		["an agreeing record, then 16 bytes before the locator", () => ({ junk: 16 })],
		["a high word on entries on this disk", () => ({ record: { 28: 1 } })],
		["a high word on total entries", () => ({ record: { 36: 1 } })],
		["a high word on the central-directory size", () => ({ record: { 44: 1 } })],
		["a high word on the central-directory offset", () => ({ record: { 52: 1 } })],
		["a high word on the record size", () => ({ record: { 8: 1 } })],
		["the locator's disk 1", () => ({ locator: { 4: 1 } })],
		["two disks", () => ({ locator: { 16: 2 } })],
		["a locator offset + 1", (c) => ({ locator: { 8: c.recordAt + 1 } })],
		["a locator offset − 1", (c) => ({ locator: { 8: c.recordAt - 1 } })],
		["a high word on the locator offset", () => ({ locator: { 12: 1 } })],
	])("rejects %s from the end-record check", async (_shape, patch) => {
		const classic = await writeZip(PARTS);
		await expectUnsupported(withZip64End(classic, patch(layoutOf(classic))), AT_END_RECORD);
	});

	// The span is checked before any record field is read: a short gap would put the record's fields
	// past the end of the buffer.
	it.each([0, 1, 30, 55])("rejects a locator after a %i-byte gap", async (keep) => {
		await expectUnsupported(withZip64End(await writeZip(PARTS), { keep }), AT_END_RECORD);
	});

	// Only the locator branch can tell this archive from undeclared trailing data: removing the
	// branch makes it corrupt-zip, and accepting any locator makes it open.
	it("reports a disagreeing record as ZIP64, not as undeclared data", async () => {
		const bytes = withZip64End(await writeZip(PARTS), { record: { 32: 3 } });
		await expectUnsupported(bytes, AT_END_RECORD);
		const noLocator = bytes.slice();
		viewOf(noLocator).setUint32(noLocator.length - EOCD_SIZE - LOCATOR_SIZE, 0, true);
		await expect(Promise.resolve().then(() => openZip(noLocator))).rejects.toMatchObject({
			code: "corrupt-zip",
		});
	});
});

describe("openZip — a ZIP64 locator inside the central directory (F11.35)", () => {
	// A reader that follows the locator sees the record's directory; a reader that ignores it sees the
	// classic one. Python's zipfile follows it, so such an archive must not open.
	it("rejects a locator in the last record's comment, with or without a ZIP64 record before it", async () => {
		const classic = await writeZip(PARTS);
		const trailer = withZip64End(classic).subarray(
			classic.length - EOCD_SIZE,
			classic.length - EOCD_SIZE + RECORD_SIZE + LOCATOR_SIZE,
		);
		const hidden = withDirectoryComment(classic, trailer);
		const view = viewOf(hidden);
		const eocd = hidden.length - EOCD_SIZE;
		expect(view.getUint32(eocd + 16, true) + view.getUint32(eocd + 12, true)).toBe(eocd);
		expect(view.getUint32(eocd - LOCATOR_SIZE, true)).toBe(0x07064b50);
		await expectUnsupported(hidden, AT_END_RECORD);
		await expectUnsupported(
			withDirectoryComment(classic, trailer.subarray(RECORD_SIZE)),
			AT_END_RECORD,
		);
	});

	it("still opens a directory whose last comment holds no locator", async () => {
		const classic = await writeZip(PARTS);
		const bytes = withDirectoryComment(classic, new Uint8Array(LOCATOR_SIZE + RECORD_SIZE));
		expect(await readAll(bytes, true)).toEqual(await readAll(classic, true));
	});

	// Fewer than 20 bytes before the end record leave no room for a locator; the guard keeps the
	// signature read inside the buffer.
	it.each([1, 19])(
		"reports %i stray bytes before an empty archive's end record as corrupt",
		async (n) => {
			const empty = await writeZip([]);
			const bytes = new Uint8Array(n + empty.length);
			bytes.set(empty, n);
			await expect(Promise.resolve().then(() => openZip(bytes))).rejects.toMatchObject({
				code: "corrupt-zip",
			});
		},
	);
});

describe("openZip — each ZIP64 rejection site has its own detail (F11.35)", () => {
	it("rejects a sentinel in the classic end record, with or without a ZIP64 end record", async () => {
		const classic = await writeZip(PARTS);
		const sentinel = classic.slice();
		viewOf(sentinel).setUint16(sentinel.length - EOCD_SIZE + 10, 0xffff, true);
		await expectUnsupported(sentinel, AT_EOCD);
		await expectUnsupported(withZip64End(sentinel), AT_EOCD);
	});

	it.each([
		["compressed size", 20],
		["uncompressed size", 24],
		["local header offset", 42],
	])(
		"rejects a sentinel %s in a central-directory entry, even after a redundant end record",
		async (_field, at) => {
			const classic = await writeZip(PARTS);
			const cdOffset = viewOf(classic).getUint32(classic.length - EOCD_SIZE + 16, true);
			const sentinel = classic.slice();
			viewOf(sentinel).setUint32(cdOffset + at, 0xffffffff, true);
			await expectUnsupported(sentinel, atCentral(cdOffset));
			await expectUnsupported(withZip64End(sentinel), atCentral(cdOffset));
		},
	);
});

describe("isRedundantZip64End (F11.35)", () => {
	async function trailer(patch: TrailerPatch = {}) {
		const bytes = withZip64End(await writeZip(PARTS), patch);
		const view = viewOf(bytes);
		const eocd = bytes.length - EOCD_SIZE;
		const classic: ClassicEnd = {
			entriesThisDisk: view.getUint16(eocd + 8, true),
			entryCount: view.getUint16(eocd + 10, true),
			cdSize: view.getUint32(eocd + 12, true),
			cdOffset: view.getUint32(eocd + 16, true),
		};
		return { view, eocd, classic, cdEnd: classic.cdOffset + classic.cdSize };
	}

	it("is true when the record and locator repeat the classic values", async () => {
		const { view, eocd, classic, cdEnd } = await trailer();
		expect(isRedundantZip64End(view, cdEnd, eocd, classic)).toBe(true);
	});

	it("is false when the gap is not one record and a locator", async () => {
		const { view, eocd, classic, cdEnd } = await trailer();
		expect(isRedundantZip64End(view, cdEnd + 1, eocd, classic)).toBe(false);
		expect(isRedundantZip64End(view, cdEnd - 1, eocd, classic)).toBe(false);
	});

	it("is false when the record disagrees", async () => {
		const { view, eocd, classic, cdEnd } = await trailer();
		expect(
			isRedundantZip64End(view, cdEnd, eocd, {
				...classic,
				entryCount: classic.entryCount + 1,
			}),
		).toBe(false);
		expect(
			isRedundantZip64End(view, cdEnd, eocd, { ...classic, cdSize: classic.cdSize - 1 }),
		).toBe(false);
	});

	it("is false when the locator disagrees", async () => {
		const { view, eocd, classic, cdEnd } = await trailer({ locator: { 16: 2 } });
		expect(isRedundantZip64End(view, cdEnd, eocd, classic)).toBe(false);
	});

	// The span is checked before any field: this view ends 30 bytes in, inside an agreeing record's
	// prefix, so reading its counts first would run past the end.
	it("checks the span before reading a record field", () => {
		const view = new DataView(new ArrayBuffer(30));
		view.setUint32(0, 0x06064b50, true);
		view.setUint32(4, 44, true);
		const classic: ClassicEnd = { entriesThisDisk: 0, entryCount: 0, cdSize: 0, cdOffset: 0 };
		expect(isRedundantZip64End(view, 0, 30, classic)).toBe(false);
	});
});

describe("smallU64 (F11.35)", () => {
	const field = (low: number, high: number) => {
		const view = new DataView(new ArrayBuffer(8));
		view.setUint32(0, low, true);
		view.setUint32(4, high, true);
		return view;
	};

	it("reads a u64 whose high word is zero as its low word", () => {
		expect(smallU64(field(44, 0), 0)).toBe(44);
		expect(smallU64(field(0xffffffff, 0), 0)).toBe(0xffffffff);
	});

	it("reads a u64 whose high word is set as -1", () => {
		expect(smallU64(field(44, 1), 0)).toBe(-1);
		expect(smallU64(field(0, 0x80000000), 0)).toBe(-1);
	});
});
