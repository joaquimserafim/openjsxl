import { XlsxError } from "../errors";
import { MAX_CLASSIC_ZIP_ENTRIES } from "./constants";
import { createEntryIntegrityVerifier, verifyEntryOutput } from "./entry-integrity";
import { inflateRaw, inflateRawStream } from "./inflate";

// Minimal reader for the ZIP (OPC) container that wraps every .xlsx. We locate the
// End-Of-Central-Directory record, walk the central directory to find each entry, and read
// its bytes on demand — inflating with ./inflate when an entry is deflate-compressed
// (method 8) or returning it verbatim when stored (method 0).
//
// Sizes and the data offset come from the CENTRAL directory + LOCAL header (never the
// local header's size fields), so the data-descriptor layout (general-purpose bit 3) that
// Excel and LibreOffice emit reads correctly.
//
// ZIP64 metadata and multi-disk archives are rejected while opening, except a ZIP64 end record that
// only repeats the classic one (F11.35). Encryption is not explicitly detected: AES method 99 is
// unsupported when read, while traditional encryption may surface as ciphertext (stored entries) or
// as an inflate/corrupt error (deflated entries).
//
// Entry policy: directory placeholders (`name/`) are skipped; duplicate part names are refused
// (OPC forbids them, and an ambiguous package is a zip-confusion vector). No two entries may share
// bytes, and no entry's data may reach the central directory (F11.52): many entries aliasing one
// small deflate stream would otherwise inflate it once per entry.

export interface ZipEntry {
	readonly name: string;
	/** 0 = stored, 8 = raw deflate */
	readonly method: number;
	/** CRC-32 of the uncompressed entry bytes, from the central directory. */
	readonly crc32: number;
	readonly compressedSize: number;
	readonly uncompressedSize: number;
	/** absolute offset of the local file header within the archive */
	readonly localHeaderOffset: number;
}

export interface ZipArchive {
	readonly entries: ReadonlyMap<string, ZipEntry>;
	has(name: string): boolean;
	/**
	 * Read an entry whole. Fails typed `missing-part`, `part-too-large` (`maxPartBytes`, the ratio cap,
	 * or the `maxTotalBytes` budget), `corrupt-zip` or `unsupported`.
	 */
	read(name: string): Promise<Uint8Array>;
	/**
	 * Read an entry as a stream of chunks, without materializing the whole part: a stored entry in
	 * 16 KiB slices, a deflated one as it inflates. It fails as `read` does; the `maxTotalBytes` budget
	 * is charged for the whole part before the first chunk, and the size and CRC-32 checks run at
	 * natural EOF only.
	 *
	 * @example for await (const chunk of zip.readStream("xl/worksheets/sheet1.xml")) total += chunk.byteLength // a 40,000-byte stored part: chunks of 16384, 16384 and 7232 bytes
	 * @example zip.readStream("nope.xml").next() // rejects XlsxError "missing-part"
	 * @example openZip(bytes, { maxTotalBytes: 10 }).readStream("xl/workbook.xml").next() // rejects XlsxError "part-too-large", before any chunk
	 * @example openZip(damaged, { verifyChecksums: true }).readStream(name) // yields every chunk, then rejects XlsxError "corrupt-zip"
	 * @example zip.readStream("aes.xml").next() // rejects XlsxError "unsupported" (compression method 99)
	 */
	readStream(name: string): AsyncGenerator<Uint8Array>;
}

const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
// A ZIP64 end record with no extensible data: 56 bytes, whose own size field counts the 44 after it.
const ZIP64_EOCD_SIZE = 56;
const ZIP64_EOCD_REMAINING = 44;
const MAX_COMMENT = 0xffff;
const ZIP64_SENTINEL = 0xffffffff;

// ── Decompression-bomb defaults (F9.7) ─────────────────────────────────────────────────────────
// A default read is safe against a hostile archive without the caller opting in. Two guards, both
// caller-tunable (raise, or set to Infinity to disable):
//   • an ABSOLUTE ceiling on any one part's output — the pathological single-huge-part backstop;
//   • a compression-RATIO cap over a floor — catches a bomb (deflate reaches ~1000:1) early, while
//     real spreadsheet XML (~10–50:1) and small high-ratio parts (under the floor) pass untouched.
/** Absolute per-part output ceiling. Default 2 GiB; a part declaring/producing more fails typed. */
export const DEFAULT_MAX_PART_BYTES = 2 * 1024 * 1024 * 1024;
/** A part may not inflate to more than this multiple of its compressed size (above the floor). */
export const DEFAULT_MAX_COMPRESSION_RATIO = 300;
/** The ratio cap never applies below this output size, so a small part with a high ratio is fine. */
export const RATIO_FLOOR_BYTES = 8 * 1024 * 1024;

/** Where an entry's local record lies: its local header through the end of its data. */
export interface EntrySpan {
	readonly name: string;
	/** Offset of the local file header. */
	readonly start: number;
	/** Offset of the entry's data, after the local header's name and extra field. */
	readonly dataStart: number;
	/** One past the entry's last data byte (`dataStart + compressedSize`). */
	readonly end: number;
}

/**
 * The span of `entry`'s local record, read from the local header at its offset, or `undefined` when
 * no local header is there. The local header's name and extra-field lengths, which may differ from
 * the central directory's, decide where the data starts.
 *
 * @example localEntrySpan(view, entry) // { name, start: 0, dataStart: 39, end: 139 } for a 9-byte name and 100 data bytes
 * @example localEntrySpan(view, { ...entry, localHeaderOffset: view.byteLength }) // undefined
 */
export function localEntrySpan(view: DataView, entry: ZipEntry): EntrySpan | undefined {
	const start = entry.localHeaderOffset;
	if (start + 30 > view.byteLength || view.getUint32(start, true) !== SIG_LOCAL) return undefined;
	const nameLen = view.getUint16(start + 26, true);
	const extraLen = view.getUint16(start + 28, true);
	const dataStart = start + 30 + nameLen + extraLen;
	return { name: entry.name, start, dataStart, end: dataStart + entry.compressedSize };
}

/** Why a set of entry spans is not a valid archive layout. */
export type SpanConflict =
	| { readonly kind: "shared-bytes"; readonly first: string; readonly second: string }
	| { readonly kind: "into-central-directory"; readonly name: string };

/**
 * The first conflict among `spans` in local-header order: a span that reaches `cdOffset`, or two
 * spans that share bytes. `undefined` when the spans are disjoint and all end at or before
 * `cdOffset`. Spans are sorted first, so the central directory may list entries in any order.
 *
 * @example findSpanConflict([a, b], cdOffset) // undefined: a ends at or before b starts, b at or before cdOffset
 * @example findSpanConflict([a, { ...b, start: a.start }], cdOffset) // { kind: "shared-bytes", first: "a", second: "b" }
 * @example findSpanConflict([a], a.end - 1) // { kind: "into-central-directory", name: "a" }
 */
export function findSpanConflict(
	spans: readonly EntrySpan[],
	cdOffset: number,
): SpanConflict | undefined {
	const sorted = [...spans].sort((x, y) => x.start - y.start);
	let previous: EntrySpan | undefined;
	for (const span of sorted) {
		if (span.end > cdOffset) return { kind: "into-central-directory", name: span.name };
		if (previous !== undefined && previous.end > span.start) {
			return { kind: "shared-bytes", first: previous.name, second: span.name };
		}
		previous = span;
	}
	return undefined;
}

/**
 * The ZIP read options `openZip` takes, as a caller passes them: the fields of `ReadOptions`, not yet
 * checked.
 */
export interface ZipReadOptions {
	readonly maxPartBytes?: number;
	readonly maxCompressionRatio?: number;
	readonly maxTotalBytes?: number;
	readonly verifyChecksums?: boolean;
}

/** The ZIP read options after {@link zipLimits}: each one read once, checked and defaulted. */
export interface ZipLimits {
	readonly maxPartBytes: number;
	readonly maxCompressionRatio: number;
	readonly maxTotalBytes: number;
	readonly verifyChecksums: boolean;
}

/**
 * Read and check a caller's ZIP read options (F11.41(a)), each property exactly once. An absent
 * option (`undefined` or `null`, or no options at all) keeps its default: `maxPartBytes` 2 GiB,
 * `maxCompressionRatio` 300, `maxTotalBytes` unlimited and `verifyChecksums` false. A limit must be a
 * number, not `NaN`, at least 0, with `Infinity` for no limit, and `verifyChecksums` must be a
 * boolean. Any other value fails typed `invalid-input`, so one such as `Number(process.env.UNSET)`
 * cannot switch a guard off. `openZip` calls it before anything else, so every reader that opens an
 * archive checks its options before any container work; `detectSpreadsheetFormat` calls it itself,
 * outside the catch that turns an unreadable archive into `undefined`.
 *
 * @example zipLimits(undefined) // { maxPartBytes: 2147483648, maxCompressionRatio: 300, maxTotalBytes: Infinity, verifyChecksums: false }
 * @example zipLimits({ maxTotalBytes: 1024 }).maxTotalBytes // 1024
 * @example zipLimits({ maxCompressionRatio: Number.POSITIVE_INFINITY }).maxCompressionRatio // Infinity (no ratio limit)
 * @example zipLimits({ maxPartBytes: Number.NaN }) // throws XlsxError "invalid-input"
 * @example zipLimits({ verifyChecksums: 1 }) // throws XlsxError "invalid-input"
 */
export function zipLimits(options: ZipReadOptions | undefined): ZipLimits {
	// Each property is read once into a local, as `unknown`: a JS caller can pass anything, and a
	// getter must not hand one value to the check and another to the reader.
	const maxPartBytes: unknown = options?.maxPartBytes;
	const maxCompressionRatio: unknown = options?.maxCompressionRatio;
	const maxTotalBytes: unknown = options?.maxTotalBytes;
	const verifyChecksums: unknown = options?.verifyChecksums;
	return {
		maxPartBytes: limitOption(
			maxPartBytes,
			DEFAULT_MAX_PART_BYTES,
			"ReadOptions.maxPartBytes must be a number of bytes, at least 0 (Infinity for no limit)",
		),
		maxCompressionRatio: limitOption(
			maxCompressionRatio,
			DEFAULT_MAX_COMPRESSION_RATIO,
			"ReadOptions.maxCompressionRatio must be a number, at least 0 (Infinity for no limit)",
		),
		maxTotalBytes: limitOption(
			maxTotalBytes,
			Number.POSITIVE_INFINITY,
			"ReadOptions.maxTotalBytes must be a number of bytes, at least 0 (Infinity for no limit)",
		),
		verifyChecksums: booleanOption(
			verifyChecksums,
			false,
			"ReadOptions.verifyChecksums must be true or false",
		),
	};
}

/**
 * One limit option, by the rule F11.53 set for `maxTotalBytes` (F11.41(a)): `fallback` when absent
 * (`undefined` or `null`), else the value, which must be a number, not `NaN`, at least 0, with
 * `Infinity` allowed. Any other value throws `XlsxError("invalid-input", message)`.
 *
 * @example limitOption(undefined, 300, message) // 300
 * @example limitOption(null, 300, message) // 300
 * @example limitOption(0, 300, message) // 0
 * @example limitOption(Number.POSITIVE_INFINITY, 300, message) // Infinity
 * @example limitOption("100", 300, message) // throws XlsxError "invalid-input" with that message
 */
export function limitOption(value: unknown, fallback: number, message: string): number {
	if (value === undefined || value === null) return fallback;
	if (typeof value !== "number" || Number.isNaN(value) || value < 0) {
		throw new XlsxError("invalid-input", message);
	}
	return value;
}

/**
 * One boolean option (F11.41(a)): `fallback` when absent (`undefined` or `null`), else the value,
 * which must be `true` or `false`. Any other value, such as `"false"` or 1, throws
 * `XlsxError("invalid-input", message)`.
 *
 * @example booleanOption(undefined, false, message) // false
 * @example booleanOption(true, false, message) // true
 * @example booleanOption("false", false, message) // throws XlsxError "invalid-input" with that message
 */
export function booleanOption(value: unknown, fallback: boolean, message: string): boolean {
	if (value === undefined || value === null) return fallback;
	if (typeof value !== "boolean") throw new XlsxError("invalid-input", message);
	return value;
}

/**
 * The tightest output cap for inflating an entry (F9.7): its declared uncompressed size, the
 * `maxPartBytes` ceiling, and the ratio limit, which is the larger of the 8 MiB floor and the
 * compressed size times `maxCompressionRatio`. A limit lowers the cap only when it is a number below
 * it, so the declared size caps the inflate whatever the options are (F11.41(a)): a `NaN` limit, or
 * an `Infinity` ratio over an empty payload (`0 × Infinity`), cannot lift it.
 *
 * @example inflateCap(100, 10, 2 ** 31, 300) // 100: the declared size
 * @example inflateCap(100, 10, 50, 300) // 50: maxPartBytes
 * @example inflateCap(2 ** 30, 1000, 2 ** 31, 300) // 8388608: the ratio floor
 * @example inflateCap(2 ** 30, 2 ** 20, 2 ** 31, 300) // 314572800: the compressed size times the ratio
 * @example inflateCap(100, 10, Number.NaN, Number.NaN) // 100: the declared size still caps
 */
export function inflateCap(
	declared: number,
	compressed: number,
	maxPartBytes: number,
	maxRatio: number,
): number {
	let cap = declared;
	if (maxPartBytes < cap) cap = maxPartBytes;
	const ratioLimit = Math.max(RATIO_FLOOR_BYTES, compressed * maxRatio);
	if (ratioLimit < cap) cap = ratioLimit;
	return cap;
}

/**
 * A meter for the aggregate inflate budget (F11.53). Its `charge(name, bytes)` counts the most that a
 * read of part `name` can produce, before the part is read, and fails typed `part-too-large` when that
 * would take the archive's total past `limit`. Without a limit it counts nothing. Under a finite limit
 * a bound that is not a number is refused too: such a part has no cap, so the budget cannot vouch for
 * it. `openZip` no longer produces one, since `zipLimits` refuses a `NaN` option (F11.41(a)).
 *
 * @example budgetMeter(Number.POSITIVE_INFINITY)("xl/a.xml", Number.NaN) // allowed: no budget
 * @example budgetMeter(100)("xl/a.xml", Number.NaN) // throws XlsxError "part-too-large" (no bound)
 * @example budgetMeter(100)("xl/a.xml", 60) // counted: 60 of 100 used
 * @example const charge = budgetMeter(100); charge("xl/a.xml", 60); charge("xl/b.xml", 41) // throws XlsxError "part-too-large"
 */
export function budgetMeter(limit: number): (name: string, bytes: number) => void {
	let used = 0;
	return (name, bytes) => {
		if (limit === Number.POSITIVE_INFINITY) return;
		if (Number.isNaN(bytes)) {
			throw new XlsxError(
				"part-too-large",
				`zip part ${name} has no size bound (maxPartBytes or maxCompressionRatio is not a number), so it cannot be read within the ${limit}-byte maxTotalBytes limit`,
			);
		}
		if (bytes > limit - used) {
			throw new XlsxError(
				"part-too-large",
				`zip part ${name} would bring the decompressed total to ${used + bytes} bytes, over the ${limit}-byte maxTotalBytes limit`,
			);
		}
		used += bytes;
	};
}

// The slice a stored part is streamed in (F11.98): the size of the chunks Node's inflate yields, so
// a stored part reaches the XML stream in the same pieces as a deflated one.
const STORED_SLICE_BYTES = 16 * 1024;

/**
 * Hand a stored part's bytes out in 16 KiB slices, the way an inflated part arrives in chunks, so a
 * stream consumer never takes a large part in one piece (F11.98). Each slice is a view into
 * `payload`; nothing is copied.
 *
 * @example [...storedSlices(new Uint8Array(0))] // [] (an empty part yields nothing)
 * @example [...storedSlices(new Uint8Array(100))].map((s) => s.byteLength) // [100]
 * @example [...storedSlices(new Uint8Array(3 * 16384 + 123))].map((s) => s.byteLength) // [16384, 16384, 16384, 123]
 */
export function* storedSlices(payload: Uint8Array): Generator<Uint8Array> {
	for (let start = 0; start < payload.byteLength; start += STORED_SLICE_BYTES) {
		yield payload.subarray(start, start + STORED_SLICE_BYTES);
	}
}

/** The classic end-of-central-directory values that a redundant ZIP64 end record repeats. */
export interface ClassicEnd {
	readonly entriesThisDisk: number;
	readonly entryCount: number;
	readonly cdSize: number;
	readonly cdOffset: number;
}

/**
 * A u64 field as a number when its high word is zero, else -1, which matches no count, size or
 * offset: a value past 2^32 - 1 is real ZIP64.
 *
 * @example smallU64(view, at) // 44 for the bytes 2c 00 00 00 00 00 00 00
 * @example smallU64(view, at) // -1 for the bytes 2c 00 00 00 01 00 00 00 (the high word is set)
 */
export function smallU64(view: DataView, at: number): number {
	return view.getUint32(at + 4, true) === 0 ? view.getUint32(at, true) : -1;
}

/**
 * Whether the ZIP64 end record and locator between the central directory, which ends at `cdEnd`, and
 * the classic end record at `eocd` only repeat the classic values (F11.35). bsdtar's `zip:zip64`, JDK
 * zipfs with `forceZIP64End` and Info-ZIP reading stdin write them. The record must fill the gap
 * exactly, checked before any field is read; it has no extensible data, sits on disk 0 and holds
 * `classic`'s counts, size and offset. The locator must name disk 0, the record's offset and at most
 * one disk. Every u64 needs a zero high word. Anything else is real ZIP64 (APPNOTE 4.3.14, 4.3.15).
 * The version fields are not checked. The caller has found the locator's signature at `eocd - 20`,
 * with `0 <= cdEnd <= eocd` and the classic end record inside the view.
 *
 * @example isRedundantZip64End(view, cdEnd, cdEnd + 76, classic) // true: bsdtar's zip:zip64 trailer
 * @example isRedundantZip64End(view, cdEnd, cdEnd + 77, classic) // false: the gap is not one record and a locator
 * @example isRedundantZip64End(view, cdEnd, eocd, { ...classic, entryCount: classic.entryCount + 1 }) // false: the record disagrees
 * @example isRedundantZip64End(twoDiskView, cdEnd, eocd, classic) // false: the locator counts two disks
 */
export function isRedundantZip64End(
	view: DataView,
	cdEnd: number,
	eocd: number,
	classic: ClassicEnd,
): boolean {
	const locator = eocd - ZIP64_LOCATOR_SIZE;
	if (locator - cdEnd !== ZIP64_EOCD_SIZE) return false;
	const recordAgrees =
		view.getUint32(cdEnd, true) === SIG_ZIP64_EOCD &&
		smallU64(view, cdEnd + 4) === ZIP64_EOCD_REMAINING &&
		view.getUint32(cdEnd + 16, true) === 0 &&
		view.getUint32(cdEnd + 20, true) === 0 &&
		smallU64(view, cdEnd + 24) === classic.entriesThisDisk &&
		smallU64(view, cdEnd + 32) === classic.entryCount &&
		smallU64(view, cdEnd + 40) === classic.cdSize &&
		smallU64(view, cdEnd + 48) === classic.cdOffset;
	if (!recordAgrees) return false;
	return (
		view.getUint32(locator + 4, true) === 0 &&
		smallU64(view, locator + 8) === cdEnd &&
		view.getUint32(locator + 16, true) <= 1
	);
}

// The EOCD record sits at the very end, after an optional ≤64 KB comment. Scan backwards
// for its signature and confirm the comment length lands exactly on the end of file (so a
// signature appearing inside compressed data can't be mistaken for it).
function findEocd(view: DataView, len: number): number {
	const earliest = Math.max(0, len - EOCD_MIN_SIZE - MAX_COMMENT);
	for (let pos = len - EOCD_MIN_SIZE; pos >= earliest; pos--) {
		if (view.getUint32(pos, true) !== SIG_EOCD) continue;
		const commentLen = view.getUint16(pos + 20, true);
		if (pos + EOCD_MIN_SIZE + commentLen === len) return pos;
	}
	return -1;
}

/**
 * Open a classic ZIP archive (no ZIP64 beyond a redundant end record, one disk) and check its
 * layout: every central-directory record, duplicate names, and entries that share bytes or reach
 * the central directory. Parts are read on demand through the returned archive. The options are the
 * zip-bomb guards of `ReadOptions`: `maxPartBytes` and `maxCompressionRatio` cap each part, and
 * `maxTotalBytes` caps every part read through this archive together (F11.53); `verifyChecksums`
 * checks each fully read part's CRC-32. A read fails typed `missing-part`, `part-too-large`,
 * `corrupt-zip` or `unsupported`.
 *
 * @example openZip(await writeZip([{ name: "a.xml", data }])).has("a.xml") // true
 * @example openZip(bytes, { maxTotalBytes: Number.NaN }) // throws XlsxError "invalid-input", before any parsing
 * @example openZip(new Uint8Array(4), { maxPartBytes: -1 }) // throws XlsxError "invalid-input", not "not-a-zip"
 * @example openZip(new Uint8Array(4)) // throws XlsxError "not-a-zip" (no end-of-central-directory record)
 * @example openZip(bsdtarZip64).has("xl/workbook.xml") // true: its ZIP64 end record only repeats the classic one
 * @example openZip(zip64) // throws XlsxError "unsupported" (ZIP64, or a multi-disk archive)
 * @example openZip(truncated) // throws XlsxError "corrupt-zip" (a bad, oversized or mismatched central directory)
 * @example openZip(withDuplicateName) // throws XlsxError "corrupt-zip" (two entries with one name)
 * @example openZip(withSharedBytes) // throws XlsxError "corrupt-zip" (two entries share bytes, or one reaches the central directory)
 */
export function openZip(bytes: Uint8Array, options?: ZipReadOptions): ZipArchive {
	// The options are read and checked first (F11.41(a)), so a bad value fails before any parsing.
	const limits = zipLimits(options);
	const charge = budgetMeter(limits.maxTotalBytes);
	const len = bytes.byteLength;
	const view = new DataView(bytes.buffer, bytes.byteOffset, len);
	// Zip-bomb guards (F9.7), applied by DEFAULT so a hostile archive is safe without opt-in.
	// Pass a larger number to raise, or Number.POSITIVE_INFINITY to disable either guard.
	const { maxPartBytes, maxCompressionRatio: maxRatio, verifyChecksums } = limits;

	const eocd = findEocd(view, len);
	if (eocd === -1) {
		throw new XlsxError(
			"not-a-zip",
			"not a zip archive: end-of-central-directory record not found",
		);
	}

	const diskNumber = view.getUint16(eocd + 4, true);
	const centralDirectoryDisk = view.getUint16(eocd + 6, true);
	const entriesThisDisk = view.getUint16(eocd + 8, true);
	const entryCount = view.getUint16(eocd + 10, true);
	const cdSize = view.getUint32(eocd + 12, true);
	const cdOffset = view.getUint32(eocd + 16, true);
	if (
		entriesThisDisk > MAX_CLASSIC_ZIP_ENTRIES ||
		entryCount > MAX_CLASSIC_ZIP_ENTRIES ||
		cdSize === ZIP64_SENTINEL ||
		cdOffset === ZIP64_SENTINEL
	) {
		throw new XlsxError(
			"unsupported",
			"ZIP64 archives are not supported: the end-of-central-directory record holds a ZIP64 sentinel",
		);
	}
	if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesThisDisk !== entryCount) {
		throw new XlsxError("unsupported", "multi-disk ZIP archives are not supported");
	}
	if (cdOffset > eocd || cdSize > eocd - cdOffset) {
		throw new XlsxError(
			"corrupt-zip",
			"corrupt zip: central directory overlaps the end-of-central-directory record",
		);
	}
	const cdEnd = cdOffset + cdSize;
	// A ZIP64 locator before the end record is checked wherever it lies, even inside the central
	// directory: a reader that follows it would see another directory (F11.35). A ZIP64 end record
	// that only repeats the classic values is redundant and read as classic; one that differs in any
	// checked field is real ZIP64.
	const zip64Locator =
		eocd >= ZIP64_LOCATOR_SIZE &&
		view.getUint32(eocd - ZIP64_LOCATOR_SIZE, true) === SIG_ZIP64_LOCATOR;
	if (zip64Locator) {
		if (
			!isRedundantZip64End(view, cdEnd, eocd, {
				entriesThisDisk,
				entryCount,
				cdSize,
				cdOffset,
			})
		) {
			throw new XlsxError(
				"unsupported",
				"ZIP64 archives are not supported: the ZIP64 end record does not repeat the end-of-central-directory record",
			);
		}
	} else if (cdEnd !== eocd) {
		throw new XlsxError(
			"corrupt-zip",
			"corrupt zip: undeclared data between the central directory and end-of-central-directory record",
		);
	}

	const decoder = new TextDecoder();
	const entries = new Map<string, ZipEntry>();
	let pos = cdOffset;
	for (let i = 0; i < entryCount; i++) {
		if (pos > cdEnd || 46 > cdEnd - pos || view.getUint32(pos, true) !== SIG_CENTRAL) {
			throw new XlsxError(
				"corrupt-zip",
				`corrupt zip: bad central directory header at offset ${pos}`,
			);
		}
		const method = view.getUint16(pos + 10, true);
		const entryCrc32 = view.getUint32(pos + 16, true);
		const compressedSize = view.getUint32(pos + 20, true);
		const uncompressedSize = view.getUint32(pos + 24, true);
		const nameLen = view.getUint16(pos + 28, true);
		const extraLen = view.getUint16(pos + 30, true);
		const commentLen = view.getUint16(pos + 32, true);
		const diskStart = view.getUint16(pos + 34, true);
		const localHeaderOffset = view.getUint32(pos + 42, true);
		if (
			compressedSize === ZIP64_SENTINEL ||
			uncompressedSize === ZIP64_SENTINEL ||
			localHeaderOffset === ZIP64_SENTINEL
		) {
			throw new XlsxError(
				"unsupported",
				`ZIP64 archives are not supported: the central directory entry at offset ${pos} holds a ZIP64 sentinel`,
			);
		}
		if (diskStart !== 0) {
			throw new XlsxError("unsupported", "multi-disk ZIP archives are not supported");
		}
		const recordSize = 46 + nameLen + extraLen + commentLen;
		if (recordSize > cdEnd - pos) {
			throw new XlsxError(
				"corrupt-zip",
				`corrupt zip: central directory entry at offset ${pos} exceeds the declared directory size`,
			);
		}
		const name = decoder.decode(bytes.subarray(pos + 46, pos + 46 + nameLen));
		pos += recordSize;
		// Directory entries (`name/`) are placeholders, not parts — skip them.
		if (name.endsWith("/")) continue;
		// OPC forbids duplicate part names; a duplicate makes the package ambiguous (a
		// zip-confusion vector), so refuse it rather than silently resolving to one entry.
		if (entries.has(name)) {
			throw new XlsxError("corrupt-zip", `corrupt zip: duplicate entry name ${name}`);
		}
		entries.set(name, {
			name,
			method,
			crc32: entryCrc32,
			compressedSize,
			uncompressedSize,
			localHeaderOffset,
		});
	}
	if (pos !== cdEnd) {
		throw new XlsxError(
			"corrupt-zip",
			"corrupt zip: central directory size does not match its entry records",
		);
	}

	// Every readable entry's span, checked before any part is read. An entry without a local header
	// has no span; it cannot be read (locate reports it), so it cannot share bytes either.
	const spans = new Map<string, EntrySpan>();
	for (const entry of entries.values()) {
		const span = localEntrySpan(view, entry);
		if (span !== undefined) spans.set(entry.name, span);
	}
	const conflict = findSpanConflict([...spans.values()], cdOffset);
	if (conflict?.kind === "shared-bytes") {
		throw new XlsxError(
			"corrupt-zip",
			`corrupt zip: entries ${conflict.first} and ${conflict.second} share bytes`,
		);
	}
	if (conflict !== undefined) {
		throw new XlsxError(
			"corrupt-zip",
			`corrupt zip: entry data for ${conflict.name} runs past the start of the central directory`,
		);
	}

	// Find an entry's raw (still-compressed) payload through the span validated above.
	function locate(name: string): { entry: ZipEntry; payload: Uint8Array } {
		const entry = entries.get(name);
		if (entry === undefined)
			throw new XlsxError("missing-part", `zip entry not found: ${name}`);
		// The part's OUTPUT size against the absolute ceiling. For a STORED entry the output is
		// the compressedSize bytes we return (uncompressedSize is untrusted and can lie small —
		// the pre-F9.7 gap that let a large stored part slip a maxPartBytes guard); for deflate it
		// is the declared uncompressedSize (the inflate re-checks the ACTUAL size while streaming).
		const declaredOutput = entry.method === 0 ? entry.compressedSize : entry.uncompressedSize;
		if (declaredOutput > maxPartBytes) {
			throw new XlsxError(
				"part-too-large",
				`zip part ${name} declares ${declaredOutput} bytes, over the ${maxPartBytes}-byte limit`,
			);
		}
		const span = spans.get(name);
		if (span === undefined) {
			throw new XlsxError("corrupt-zip", `corrupt zip: bad local header for ${name}`);
		}
		return { entry, payload: bytes.subarray(span.dataStart, span.end) };
	}

	// The tightest output cap for inflating `entry` (inflateCap). The inflate aborts typed as soon as
	// the ACTUAL output would exceed it, so a bomb that lies about its size can't expand.
	function outputCap(entry: ZipEntry): number {
		return inflateCap(entry.uncompressedSize, entry.compressedSize, maxPartBytes, maxRatio);
	}

	// The most a read of `entry` can produce, charged to the aggregate budget before it is read: a
	// stored part's own bytes, or a deflated part's inflate cap, past which the inflate aborts.
	function outputBound(entry: ZipEntry, payload: Uint8Array): number {
		if (entry.method === 0) return payload.byteLength;
		return entry.compressedSize === 0 ? 0 : outputCap(entry);
	}

	async function read(name: string): Promise<Uint8Array> {
		const { entry, payload } = locate(name);
		if (entry.method === 0) {
			charge(name, outputBound(entry, payload));
			verifyEntryOutput(entry, payload, verifyChecksums);
			return payload;
		}
		if (entry.method === 8) {
			charge(name, outputBound(entry, payload));
			const output =
				entry.compressedSize === 0
					? new Uint8Array(0)
					: await inflateRaw(payload, outputCap(entry), name);
			verifyEntryOutput(entry, output, verifyChecksums);
			return output;
		}
		throw new XlsxError(
			"unsupported",
			`unsupported zip compression method ${entry.method} for ${name}`,
		);
	}

	// Read an entry as a stream of chunks, never materializing the whole part: a stored part in
	// 16 KiB slices, a deflated one as it inflates. The budget is charged for the whole
	// part before the first chunk. Size/CRC checks run only after natural EOF, so a caller that
	// cancels early makes no whole-entry integrity claim.
	async function* readStream(name: string): AsyncGenerator<Uint8Array> {
		const { entry, payload } = locate(name);
		const integrity = createEntryIntegrityVerifier(entry, verifyChecksums);
		if (entry.method === 0) {
			charge(name, outputBound(entry, payload));
			for (const slice of storedSlices(payload)) {
				integrity.update(slice);
				yield slice;
			}
			integrity.finish();
			return;
		}
		if (entry.method === 8) {
			charge(name, outputBound(entry, payload));
			if (entry.compressedSize > 0) {
				for await (const chunk of inflateRawStream(payload, outputCap(entry), name)) {
					integrity.update(chunk);
					yield chunk;
				}
			}
			integrity.finish();
			return;
		}
		throw new XlsxError(
			"unsupported",
			`unsupported zip compression method ${entry.method} for ${name}`,
		);
	}

	return {
		entries,
		has: (name) => entries.has(name),
		read,
		readStream,
	};
}
