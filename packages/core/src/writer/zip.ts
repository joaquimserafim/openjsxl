import { XlsxError } from "../errors"
import { crc32 } from "./crc32"
import { deflateRaw } from "./deflate"

// Write the OPC/ZIP container that wraps every .xlsx — the mirror image of zip/central-directory.ts.
// Given a list of named byte-parts, emit an archive whose bytes re-read byte-identically through
// `openZip`: a run of [local header + payload] blocks, then the central directory, then the
// end-of-central-directory (EOCD) record. Every multi-byte field is little-endian.
//
// Compression. Each part is deflated (method 8) and the result kept only when it is actually
// smaller than the raw bytes; otherwise the part is stored (method 0). A 200-byte `.rels` or an
// already-compressed blob shouldn't pay a deflate that makes it *bigger*, and "store when it
// doesn't help" is what every real zip writer does. We buffer each part in full, so the CRC and
// both sizes are known before the header is written and go straight into it — no data descriptor
// (general-purpose bit 3 stays 0), the simplest layout the reader accepts.
//
// Determinism. A fixed DOS timestamp (no clock read) means identical parts produce byte-identical
// archives, so golden-file tests can assert on the exact bytes.
//
// Limits. Classic ZIP caps sizes/offsets at a u32 and the entry count at a u16; the values
// 0xFFFFFFFF / 0xFFFF are ZIP64 sentinels. The reader rejects ZIP64, so the writer refuses to
// produce anything that would need it — it never emits a sentinel that would be misread. These
// ceilings are ~4 GB and 65 534 parts, unreachable by any real spreadsheet.
//
// This layer knows nothing about worksheets; it just packs named parts. The OOXML wiring is F3.2.

const encoder = new TextEncoder()

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50

const METHOD_STORE = 0
const METHOD_DEFLATE = 8
// Version 2.0 — the ZIP spec level that introduced DEFLATE. Used for "version needed to extract"
// (and "version made by") on every entry, whether stored or deflated, to keep the layout uniform.
const VERSION = 20

// u32 size/offset fields must stay strictly below the ZIP64 sentinel; the u16 entry count strictly
// below its own. Hitting either forces ZIP64, which we don't emit.
const U32_CEILING = 0xffffffff
const MAX_ENTRIES = 0xffff

// Fixed DOS date (1980-01-01) and time (00:00) for deterministic output — mirrors the fixtures
// builder so both produce reproducible archives.
const DOS_TIME = 0
const DOS_DATE = 0x0021

const u16 = (n: number): Uint8Array => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff])
const u32 = (n: number): Uint8Array =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff])

function concat(parts: Uint8Array[]): Uint8Array {
	let total = 0
	for (const part of parts) total += part.length
	const out = new Uint8Array(total)
	let offset = 0
	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}
	return out
}

/** One part to pack: an OPC path (e.g. `xl/workbook.xml`) and its uncompressed bytes. */
export interface ZipInput {
	readonly name: string
	readonly data: Uint8Array
}

/**
 * Pack named byte-parts into a ZIP/OPC archive that `openZip` reads back byte-for-byte. Async
 * because compression runs on the platform's CompressionStream.
 *
 * Throws {@link XlsxError} with code `unsupported` when the archive would exceed a classic-ZIP
 * field limit — too many entries, a part too large, or a name longer than the u16 name-length
 * field (all of which would otherwise require ZIP64, which the reader rejects). Throws a plain
 * `Error` on a caller contract violation the reader would silently mishandle: a duplicate name, or
 * a name ending in `/` (which the reader treats as a directory placeholder and drops).
 */
export async function writeZip(entries: readonly ZipInput[]): Promise<Uint8Array> {
	if (entries.length >= MAX_ENTRIES) {
		throw new XlsxError(
			"unsupported",
			`too many zip entries (${entries.length}); would require ZIP64, which is not supported`,
		)
	}

	const seen = new Set<string>()
	const local: Uint8Array[] = []
	const central: Uint8Array[] = []
	let offset = 0

	for (const entry of entries) {
		// A name ending in `/` is a ZIP directory placeholder: the reader skips it (line 108 of
		// central-directory.ts) so it never enters the entries map. Writing a real part under such
		// a name would emit bytes that vanish on read — refuse it rather than break the round-trip.
		if (entry.name.endsWith("/")) {
			throw new Error(`invalid zip entry name (directory placeholder): ${entry.name}`)
		}
		if (seen.has(entry.name)) throw new Error(`duplicate zip entry name: ${entry.name}`)
		seen.add(entry.name)

		const name = encoder.encode(entry.name)
		// The name-length header field is a u16 with no ZIP64 escape. A name whose UTF-8 encoding
		// exceeds 65535 bytes would be silently truncated by u16() while the full name bytes are
		// still appended, desyncing the reader's payload offset (header+30+nameLen+extraLen) and
		// corrupting the archive. Guard it like the other classic-ZIP ceilings.
		if (name.length > 0xffff) {
			throw new XlsxError(
				"unsupported",
				`zip entry name too long (${name.length} bytes); exceeds the classic-ZIP name-length field`,
			)
		}
		const data = entry.data
		const crc = crc32(data)
		const uncompressedSize = data.length

		// Compress, but keep it only when it's a strict win — otherwise store the raw bytes.
		const deflated = await deflateRaw(data)
		const useDeflate = deflated.length < data.length
		const method = useDeflate ? METHOD_DEFLATE : METHOD_STORE
		const payload = useDeflate ? deflated : data
		const compressedSize = payload.length

		if (
			uncompressedSize >= U32_CEILING ||
			compressedSize >= U32_CEILING ||
			offset >= U32_CEILING
		) {
			throw new XlsxError(
				"unsupported",
				`zip entry ${entry.name} too large for a classic (ZIP64-free) archive`,
			)
		}

		const header = concat([
			u32(SIG_LOCAL),
			u16(VERSION), // version needed to extract
			u16(0), // general-purpose flags: none (sizes/CRC are in the header, no data descriptor)
			u16(method),
			u16(DOS_TIME),
			u16(DOS_DATE),
			u32(crc),
			u32(compressedSize),
			u32(uncompressedSize),
			u16(name.length),
			u16(0), // extra field length
			name,
		])
		local.push(header, payload)

		central.push(
			concat([
				u32(SIG_CENTRAL),
				u16(VERSION), // version made by
				u16(VERSION), // version needed to extract
				u16(0), // flags
				u16(method),
				u16(DOS_TIME),
				u16(DOS_DATE),
				u32(crc),
				u32(compressedSize),
				u32(uncompressedSize),
				u16(name.length),
				u16(0), // extra field length
				u16(0), // file comment length
				u16(0), // disk number start
				u16(0), // internal file attributes
				u32(0), // external file attributes
				u32(offset), // relative offset of the local header
				name,
			]),
		)

		offset += header.length + payload.length
	}

	const directory = concat(central)
	// After the loop, `offset` is the total size of all local blocks — i.e. where the central
	// directory begins. Guard it and the directory size against the same u32 ceiling.
	if (offset >= U32_CEILING || directory.length >= U32_CEILING) {
		throw new XlsxError(
			"unsupported",
			"zip archive too large for a classic (ZIP64-free) archive",
		)
	}

	const eocd = concat([
		u32(SIG_EOCD),
		u16(0), // number of this disk
		u16(0), // disk where the central directory starts
		u16(entries.length), // central-directory entries on this disk
		u16(entries.length), // total central-directory entries
		u32(directory.length), // size of the central directory
		u32(offset), // offset of the central directory from the start of the archive
		u16(0), // comment length
	])

	return concat([...local, directory, eocd])
}
