import { XlsxError } from "../errors";
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
// Not supported (the reader detects and rejects these rather than misreading): ZIP64
// (> 4 GB or > 65535 entries), encryption, and multi-disk archives.
//
// Entry policy: directory placeholders (`name/`) are skipped; duplicate part names are refused
// (OPC forbids them, and an ambiguous package is a zip-confusion vector).

export interface ZipEntry {
	readonly name: string;
	/** 0 = stored, 8 = raw deflate */
	readonly method: number;
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
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const MAX_COMMENT = 0xffff;
const ZIP64_SENTINEL = 0xffffffff;

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

export function openZip(bytes: Uint8Array, options?: { maxPartBytes?: number }): ZipArchive {
	const len = bytes.byteLength;
	const view = new DataView(bytes.buffer, bytes.byteOffset, len);
	// An absolute ceiling on a single part's declared decompressed size, independent of the
	// (attacker-controllable) uncompressedSize the inflate already caps at — a zip-bomb guard
	// callers opt into. undefined ⇒ no ceiling beyond the declared size.
	const maxPartBytes = options?.maxPartBytes;

	const eocd = findEocd(view, len);
	if (eocd === -1) {
		throw new XlsxError(
			"not-a-zip",
			"not a zip archive: end-of-central-directory record not found",
		);
	}

	const entryCount = view.getUint16(eocd + 10, true);
	const cdOffset = view.getUint32(eocd + 16, true);
	if (cdOffset === ZIP64_SENTINEL)
		throw new XlsxError("unsupported", "ZIP64 archives are not supported");
	if (cdOffset > len) {
		throw new XlsxError(
			"corrupt-zip",
			"corrupt zip: central directory offset past end of file",
		);
	}

	const decoder = new TextDecoder();
	const entries = new Map<string, ZipEntry>();
	let pos = cdOffset;
	for (let i = 0; i < entryCount; i++) {
		if (pos + 46 > len || view.getUint32(pos, true) !== SIG_CENTRAL) {
			throw new XlsxError(
				"corrupt-zip",
				`corrupt zip: bad central directory header at offset ${pos}`,
			);
		}
		const method = view.getUint16(pos + 10, true);
		const compressedSize = view.getUint32(pos + 20, true);
		const uncompressedSize = view.getUint32(pos + 24, true);
		const nameLen = view.getUint16(pos + 28, true);
		const extraLen = view.getUint16(pos + 30, true);
		const commentLen = view.getUint16(pos + 32, true);
		const localHeaderOffset = view.getUint32(pos + 42, true);
		if (
			compressedSize === ZIP64_SENTINEL ||
			uncompressedSize === ZIP64_SENTINEL ||
			localHeaderOffset === ZIP64_SENTINEL
		) {
			throw new XlsxError("unsupported", "ZIP64 archives are not supported");
		}
		const name = decoder.decode(bytes.subarray(pos + 46, pos + 46 + nameLen));
		pos += 46 + nameLen + extraLen + commentLen;
		// Directory entries (`name/`) are placeholders, not parts — skip them.
		if (name.endsWith("/")) continue;
		// OPC forbids duplicate part names; a duplicate makes the package ambiguous (a
		// zip-confusion vector), so refuse it rather than silently resolving to one entry.
		if (entries.has(name)) {
			throw new XlsxError("corrupt-zip", `corrupt zip: duplicate entry name ${name}`);
		}
		entries.set(name, { name, method, compressedSize, uncompressedSize, localHeaderOffset });
	}

	// Find an entry's raw (still-compressed) payload. The local header repeats name/extra
	// lengths, which may differ from the central directory's — trust the local header to find
	// where the data actually starts.
	function locate(name: string): { entry: ZipEntry; payload: Uint8Array } {
		const entry = entries.get(name);
		if (entry === undefined)
			throw new XlsxError("missing-part", `zip entry not found: ${name}`);
		if (maxPartBytes !== undefined && entry.uncompressedSize > maxPartBytes) {
			throw new XlsxError(
				"part-too-large",
				`zip part ${name} declares ${entry.uncompressedSize} bytes, over the ${maxPartBytes}-byte limit`,
			);
		}
		const header = entry.localHeaderOffset;
		if (header + 30 > len || view.getUint32(header, true) !== SIG_LOCAL) {
			throw new XlsxError("corrupt-zip", `corrupt zip: bad local header for ${name}`);
		}
		const nameLen = view.getUint16(header + 26, true);
		const extraLen = view.getUint16(header + 28, true);
		const dataStart = header + 30 + nameLen + extraLen;
		if (dataStart + entry.compressedSize > len) {
			throw new XlsxError(
				"corrupt-zip",
				`corrupt zip: entry data for ${name} runs past end of file`,
			);
		}
		return { entry, payload: bytes.subarray(dataStart, dataStart + entry.compressedSize) };
	}

	async function read(name: string): Promise<Uint8Array> {
		const { entry, payload } = locate(name);
		if (entry.method === 0) return payload;
		if (entry.method === 8) {
			if (entry.compressedSize === 0) return new Uint8Array(0);
			try {
				return await inflateRaw(payload, entry.uncompressedSize);
			} catch (cause) {
				throw new XlsxError("corrupt-zip", `corrupt zip: failed to inflate ${name}`, {
					cause,
				});
			}
		}
		throw new XlsxError(
			"unsupported",
			`unsupported zip compression method ${entry.method} for ${name}`,
		);
	}

	// Read an entry as a stream of chunks, never materializing the whole part. Stored entries
	// yield their (already in-memory) payload as one chunk; deflate entries stream the inflate.
	async function* readStream(name: string): AsyncGenerator<Uint8Array> {
		const { entry, payload } = locate(name);
		if (entry.method === 0) {
			if (payload.byteLength > 0) yield payload;
			return;
		}
		if (entry.method === 8) {
			if (entry.compressedSize === 0) return;
			try {
				yield* inflateRawStream(payload, entry.uncompressedSize);
			} catch (cause) {
				throw new XlsxError("corrupt-zip", `corrupt zip: failed to inflate ${name}`, {
					cause,
				});
			}
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
