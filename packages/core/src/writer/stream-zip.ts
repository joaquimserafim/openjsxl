import { XlsxError } from "../errors";
import { CRC32_INIT, crc32Final, crc32Update } from "./crc32";

// Streaming ZIP/OPC writer (F5.1) — the constant-memory sibling of writeZip. Where writeZip buffers
// each part to learn its CRC and sizes before the local header, this emits an entry as it flows: a
// local header with general-purpose bit 3 set (CRC + sizes deferred), the DEFLATE payload streamed
// straight out of a CompressionStream, then a data descriptor carrying the CRC and both sizes. The
// central directory — one small record per entry — is buffered and written last with the true values
// and each entry's local-header offset, so a central-directory reader (openZip, Excel, `unzip`)
// never needs the zeroed local fields. Method is always DEFLATE: the store-if-smaller trick writeZip
// uses needs both sizes up front, which streaming forgoes. Determinism holds via the same fixed DOS
// date as writeZip; byte-identity with writeZip is explicitly NOT a goal (the layout differs by
// design) — equivalence is asserted through the reader.

const encoder = new TextEncoder();

const SIG_LOCAL = 0x04034b50;
const SIG_DATA_DESC = 0x08074b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const METHOD_DEFLATE = 8;
const VERSION = 20;
// General-purpose bit 3: the CRC-32 and the two sizes are 0 in the local header and instead follow
// the compressed data in a descriptor record. This is what lets us write the header before we know them.
const FLAG_DATA_DESCRIPTOR = 0x0008;
const U32_CEILING = 0xffffffff;
const MAX_ENTRIES = 0xffff;
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

const u16 = (n: number): Uint8Array => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff]);
const u32 = (n: number): Uint8Array =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]);

function concat(parts: Uint8Array[]): Uint8Array {
	let total = 0;
	for (const part of parts) total += part.length;
	const out = new Uint8Array(total);
	let offset = 0;
	for (const part of parts) {
		out.set(part, offset);
		offset += part.length;
	}
	return out;
}

/** A part to stream: an OPC path plus its bytes — one buffer, or a (possibly async) chunk stream. */
export interface StreamPart {
	readonly name: string;
	readonly data: Uint8Array | Iterable<Uint8Array> | AsyncIterable<Uint8Array>;
}

// Normalize any accepted data shape to an async iterable of chunks.
async function* toChunks(data: StreamPart["data"]): AsyncGenerator<Uint8Array> {
	if (data instanceof Uint8Array) {
		yield data;
		return;
	}
	for await (const chunk of data as AsyncIterable<Uint8Array>) yield chunk;
}

// DEFLATE `input` chunk-by-chunk, yielding compressed chunks while `onInput` folds each raw chunk
// into the running CRC and uncompressed-size accounting. The compressor's writable and readable run
// concurrently: a detached pump feeds input (awaiting the compressor's backpressure) while we drain
// the compressed output — so peak memory tracks the compressor's buffer, not the whole part.
async function* deflateChunks(
	input: AsyncIterable<Uint8Array>,
	onInput: (chunk: Uint8Array) => void,
): AsyncGenerator<Uint8Array> {
	const cs = new CompressionStream("deflate-raw");
	const writer = cs.writable.getWriter();
	// Drive the source through an EXPLICIT iterator so early teardown can `.return()` it — running the
	// source's `finally` (e.g. closing a DB cursor) even when the pump is parked at the source.
	const source = input[Symbol.asyncIterator]();
	const pump = (async () => {
		try {
			for (;;) {
				const next = await source.next();
				if (next.done === true) break;
				onInput(next.value);
				// The cast sidesteps a TS lib disagreement over ArrayBuffer vs SharedArrayBuffer
				// backing (same issue deflate.ts works around) — the writer accepts the bytes at runtime.
				await writer.write(next.value as Uint8Array<ArrayBuffer>);
			}
			await writer.close();
		} catch (err) {
			// Error the compressor so the drain loop below rejects with this producer error rather
			// than hanging on a writable that never closed.
			await writer.abort(err);
			throw err;
		}
	})();
	pump.catch(() => {}); // the drain loop re-throws; keep this from being an unhandled rejection
	const reader = cs.readable.getReader();
	try {
		for (;;) {
			const { done, value } = await reader.read();
			if (done) break;
			yield value;
		}
		await pump; // surface a producer error now that the output is fully drained
	} finally {
		// On ANY exit — normal end, a producer error, or an early consumer cancel that `.return()`s
		// this generator — tear everything down so the row source releases and the compressor is not
		// leaked. Cancelling the readable errors the writable, which unblocks a backpressured
		// `writer.write()`; returning the source runs its `finally` (close the cursor) even if the
		// pump is parked at `source.next()`; then the pump settles. On the normal path these are all
		// no-ops (readable drained, source exhausted, pump resolved).
		await reader.cancel().catch(() => {});
		await source.return?.().catch(() => {});
		await pump.catch(() => {});
	}
}

/**
 * Stream named parts into a ZIP/OPC archive as a chunk generator. Each part's `data` may be a single
 * buffer or a (possibly async) stream of chunks — the latter is how a worksheet's rows flow through
 * without materializing the sheet. The archive reads back through {@link openZip} identically to a
 * buffered one. Throws {@link XlsxError} `unsupported` at the classic-ZIP field ceilings (~4 GB /
 * 65 534 entries) and a plain `Error` on a duplicate or directory-placeholder name.
 */
export async function* streamZip(parts: AsyncIterable<StreamPart>): AsyncGenerator<Uint8Array> {
	const seen = new Set<string>();
	const central: Uint8Array[] = [];
	let offset = 0;
	let count = 0;

	for await (const part of parts) {
		if (part.name.endsWith("/")) {
			throw new Error(`invalid zip entry name (directory placeholder): ${part.name}`);
		}
		if (seen.has(part.name)) throw new Error(`duplicate zip entry name: ${part.name}`);
		seen.add(part.name);
		if (count >= MAX_ENTRIES) {
			throw new XlsxError(
				"unsupported",
				`too many zip entries (${count + 1}); would require ZIP64, which is not supported`,
			);
		}
		const name = encoder.encode(part.name);
		if (name.length > 0xffff) {
			throw new XlsxError(
				"unsupported",
				`zip entry name too long (${name.length} bytes); exceeds the classic-ZIP name-length field`,
			);
		}
		const localOffset = offset;
		if (localOffset >= U32_CEILING) {
			throw new XlsxError(
				"unsupported",
				"zip archive too large for a classic (ZIP64-free) archive",
			);
		}

		const header = concat([
			u32(SIG_LOCAL),
			u16(VERSION),
			u16(FLAG_DATA_DESCRIPTOR),
			u16(METHOD_DEFLATE),
			u16(DOS_TIME),
			u16(DOS_DATE),
			u32(0), // crc — in the data descriptor
			u32(0), // compressed size — in the data descriptor
			u32(0), // uncompressed size — in the data descriptor
			u16(name.length),
			u16(0), // extra field length
			name,
		]);
		yield header;
		offset += header.length;

		let crc = CRC32_INIT;
		let uncompressed = 0;
		let compressed = 0;
		for await (const chunk of deflateChunks(toChunks(part.data), (raw) => {
			crc = crc32Update(crc, raw);
			uncompressed += raw.length;
		})) {
			compressed += chunk.length;
			yield chunk;
			offset += chunk.length;
		}
		crc = crc32Final(crc);

		if (compressed >= U32_CEILING || uncompressed >= U32_CEILING || offset >= U32_CEILING) {
			throw new XlsxError(
				"unsupported",
				`zip entry ${part.name} too large for a classic (ZIP64-free) archive`,
			);
		}

		const descriptor = concat([
			u32(SIG_DATA_DESC),
			u32(crc),
			u32(compressed),
			u32(uncompressed),
		]);
		yield descriptor;
		offset += descriptor.length;

		central.push(
			concat([
				u32(SIG_CENTRAL),
				u16(VERSION), // version made by
				u16(VERSION), // version needed to extract
				u16(FLAG_DATA_DESCRIPTOR), // must match the local flags
				u16(METHOD_DEFLATE),
				u16(DOS_TIME),
				u16(DOS_DATE),
				u32(crc), // the TRUE values — a central-directory reader keys off these
				u32(compressed),
				u32(uncompressed),
				u16(name.length),
				u16(0), // extra field length
				u16(0), // file comment length
				u16(0), // disk number start
				u16(0), // internal file attributes
				u32(0), // external file attributes
				u32(localOffset), // relative offset of the local header
				name,
			]),
		);
		count++;
	}

	const directory = concat(central);
	if (offset >= U32_CEILING || directory.length >= U32_CEILING) {
		throw new XlsxError(
			"unsupported",
			"zip archive too large for a classic (ZIP64-free) archive",
		);
	}
	yield directory;
	yield concat([
		u32(SIG_EOCD),
		u16(0), // number of this disk
		u16(0), // disk where the central directory starts
		u16(count), // central-directory entries on this disk
		u16(count), // total central-directory entries
		u32(directory.length),
		u32(offset), // offset of the central directory from the start of the archive
		u16(0), // comment length
	]);
}
