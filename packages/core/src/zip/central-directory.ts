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
// ZIP64 metadata and multi-disk archives are rejected while opening. Encryption is not explicitly
// detected: AES method 99 is unsupported when read, while traditional encryption may surface as
// ciphertext (stored entries) or as an inflate/corrupt error (deflated entries).
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
	read(name: string): Promise<Uint8Array>;
	/** Read an entry as a stream of chunks, without materializing the whole part. */
	readStream(name: string): AsyncGenerator<Uint8Array>;
}

const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
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

export function openZip(
	bytes: Uint8Array,
	options?: {
		maxPartBytes?: number;
		maxCompressionRatio?: number;
		verifyChecksums?: boolean;
	},
): ZipArchive {
	const len = bytes.byteLength;
	const view = new DataView(bytes.buffer, bytes.byteOffset, len);
	// Zip-bomb guards (F9.7), applied by DEFAULT so a hostile archive is safe without opt-in.
	// Pass a larger number to raise, or Number.POSITIVE_INFINITY to disable either guard.
	const maxPartBytes = options?.maxPartBytes ?? DEFAULT_MAX_PART_BYTES;
	const maxRatio = options?.maxCompressionRatio ?? DEFAULT_MAX_COMPRESSION_RATIO;
	const verifyChecksums = options?.verifyChecksums ?? false;

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
		throw new XlsxError("unsupported", "ZIP64 archives are not supported");
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
	if (cdEnd !== eocd) {
		if (
			eocd >= ZIP64_LOCATOR_SIZE &&
			view.getUint32(eocd - ZIP64_LOCATOR_SIZE, true) === SIG_ZIP64_LOCATOR
		) {
			throw new XlsxError("unsupported", "ZIP64 archives are not supported");
		}
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
			throw new XlsxError("unsupported", "ZIP64 archives are not supported");
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

	// The tightest output cap for inflating `entry`: the declared size, the absolute ceiling, and
	// the ratio limit (max of the floor and compressedSize × ratio). The inflate aborts typed as
	// soon as the ACTUAL output would exceed it, so a bomb that lies about its size can't expand.
	function inflateCap(entry: ZipEntry): number {
		const ratioLimit = Math.max(RATIO_FLOOR_BYTES, entry.compressedSize * maxRatio);
		return Math.min(entry.uncompressedSize, maxPartBytes, ratioLimit);
	}

	async function read(name: string): Promise<Uint8Array> {
		const { entry, payload } = locate(name);
		if (entry.method === 0) {
			verifyEntryOutput(entry, payload, verifyChecksums);
			return payload;
		}
		if (entry.method === 8) {
			const output =
				entry.compressedSize === 0
					? new Uint8Array(0)
					: await inflateRaw(payload, inflateCap(entry), name);
			verifyEntryOutput(entry, output, verifyChecksums);
			return output;
		}
		throw new XlsxError(
			"unsupported",
			`unsupported zip compression method ${entry.method} for ${name}`,
		);
	}

	// Read an entry as a stream of chunks, never materializing the whole part. Size/CRC checks run
	// only after natural EOF, so a caller that cancels early makes no whole-entry integrity claim.
	async function* readStream(name: string): AsyncGenerator<Uint8Array> {
		const { entry, payload } = locate(name);
		const integrity = createEntryIntegrityVerifier(entry, verifyChecksums);
		if (entry.method === 0) {
			if (payload.byteLength > 0) {
				integrity.update(payload);
				yield payload;
			}
			integrity.finish();
			return;
		}
		if (entry.method === 8) {
			if (entry.compressedSize > 0) {
				for await (const chunk of inflateRawStream(payload, inflateCap(entry), name)) {
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
