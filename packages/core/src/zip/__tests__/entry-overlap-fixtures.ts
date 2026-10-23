// F11.52 fixtures: read and patch the central directory of a classic ZIP, to build archives whose
// entries share bytes or whose data runs into the central directory.

const SIG_EOCD = 0x06054b50;
const EOCD_SIZE = 22;
const CENTRAL_HEADER_SIZE = 46;
const LOCAL_HEADER_SIZE = 30;

/** One central-directory record, with where its entry's data starts according to its local header. */
export interface CentralRecord {
	readonly name: string;
	/** Offset of the record within the archive. */
	readonly offset: number;
	/** Length of the record: fixed header, name, extra field and comment. */
	readonly length: number;
	readonly localHeaderOffset: number;
	readonly compressedSize: number;
	readonly dataStart: number;
}

/** A classic ZIP's central directory: its offset and its records in directory order. */
export interface CentralDirectory {
	readonly offset: number;
	readonly records: readonly CentralRecord[];
}

/**
 * The central directory of a classic ZIP (no ZIP64), found through the end-of-central-directory
 * record.
 *
 * @example readCentralDirectory(await writeZip([{ name: "a", data: bytes }])).records[0]?.name // "a"
 * @example readCentralDirectory(new Uint8Array(4)) // throws Error: no end-of-central-directory record
 */
export function readCentralDirectory(bytes: Uint8Array): CentralDirectory {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let eocd = -1;
	for (let pos = bytes.byteLength - EOCD_SIZE; pos >= 0 && eocd === -1; pos--) {
		const commentLength = view.getUint16(pos + 20, true);
		if (
			view.getUint32(pos, true) === SIG_EOCD &&
			pos + EOCD_SIZE + commentLength === bytes.byteLength
		)
			eocd = pos;
	}
	if (eocd === -1) throw new Error("no end-of-central-directory record");
	const offset = view.getUint32(eocd + 16, true);
	const records: CentralRecord[] = [];
	let pos = offset;
	for (let i = view.getUint16(eocd + 10, true); i > 0; i--) {
		const nameLength = view.getUint16(pos + 28, true);
		const extraLength = view.getUint16(pos + 30, true);
		const commentLength = view.getUint16(pos + 32, true);
		const localHeaderOffset = view.getUint32(pos + 42, true);
		const localNameLength = view.getUint16(localHeaderOffset + 26, true);
		const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
		const length = CENTRAL_HEADER_SIZE + nameLength + extraLength + commentLength;
		records.push({
			name: new TextDecoder().decode(
				bytes.subarray(pos + CENTRAL_HEADER_SIZE, pos + CENTRAL_HEADER_SIZE + nameLength),
			),
			offset: pos,
			length,
			localHeaderOffset,
			compressedSize: view.getUint32(pos + 20, true),
			dataStart: localHeaderOffset + LOCAL_HEADER_SIZE + localNameLength + localExtraLength,
		});
		pos += length;
	}
	return { offset, records };
}

/**
 * The central-directory record named `name`.
 *
 * @example centralRecord(zip, "a").name // "a"
 * @example centralRecord(zip, "missing") // throws Error: no central record named missing
 */
export function centralRecord(bytes: Uint8Array, name: string): CentralRecord {
	const record = readCentralDirectory(bytes).records.find((r) => r.name === name);
	if (record === undefined) throw new Error(`no central record named ${name}`);
	return record;
}

/**
 * A copy of `bytes` whose central entry `alias` points at `source`'s local header and carries its
 * method, CRC and sizes. A reader that does not check for shared bytes reads `alias` as an exact
 * copy of `source`.
 *
 * @example aliasEntry(zip, "b", "a") // "b" now shares "a"'s local header and data
 */
export function aliasEntry(bytes: Uint8Array, alias: string, source: string): Uint8Array {
	const out = bytes.slice();
	const view = new DataView(out.buffer);
	const from = centralRecord(bytes, source).offset;
	const to = centralRecord(bytes, alias).offset;
	view.setUint16(to + 10, view.getUint16(from + 10, true), true); // method
	// CRC-32, compressed size, uncompressed size and local header offset.
	for (const field of [16, 20, 24, 42]) {
		view.setUint32(to + field, view.getUint32(from + field, true), true);
	}
	return out;
}

/**
 * A copy of `bytes` whose central entry `name` declares `size` compressed bytes.
 *
 * @example centralRecord(withCompressedSize(zip, "a", 7), "a").compressedSize // 7
 */
export function withCompressedSize(bytes: Uint8Array, name: string, size: number): Uint8Array {
	const out = bytes.slice();
	new DataView(out.buffer).setUint32(centralRecord(bytes, name).offset + 20, size, true);
	return out;
}

/**
 * A copy of `bytes` whose central entry `name` points at `offset` for its local header.
 *
 * @example centralRecord(withLocalHeaderOffset(zip, "a", 1), "a").localHeaderOffset // 1
 */
export function withLocalHeaderOffset(bytes: Uint8Array, name: string, offset: number): Uint8Array {
	const out = bytes.slice();
	new DataView(out.buffer).setUint32(centralRecord(bytes, name).offset + 42, offset, true);
	return out;
}

/**
 * A copy of `bytes` with its central-directory records in reverse order, so the directory no
 * longer lists entries in local-header order. Sizes and offsets are unchanged.
 *
 * @example readCentralDirectory(reverseCentralDirectory(zip)).records.map((r) => r.name) // ["b", "a"]
 */
export function reverseCentralDirectory(bytes: Uint8Array): Uint8Array {
	const { offset, records } = readCentralDirectory(bytes);
	const out = bytes.slice();
	let pos = offset;
	for (const record of [...records].reverse()) {
		out.set(bytes.subarray(record.offset, record.offset + record.length), pos);
		pos += record.length;
	}
	return out;
}
