// The BIFF12 binary record layer, shared by the .xlsb parsers (M7). A BIFF12 stream is a flat run
// of records, each framed as: a variable-width record id, a 7-bit varint payload length, then that
// many payload bytes. This mirrors the SAX tokenizer's role for XML — a bounded, never-throwing
// scanner the higher layers walk. Everything is little-endian; strings are UTF-16LE.
//
// Robustness (the hostile-input contract): a truncated or lying length can never read out of
// bounds — every field read is clamped to the record's own bytes, and a record whose declared
// length overruns the buffer is trimmed to what remains. An unknown record id is simply yielded and
// skipped by the caller. The framing loop always advances, so it can't hang.

const utf16 = new TextDecoder("utf-16le");

/** One framed BIFF12 record: its id and a view over just its payload bytes. */
export interface BiffRecord {
	readonly id: number;
	readonly data: Uint8Array;
}

// Read the record id: 1–2 bytes (up to 4), each carrying 8 value bits, continuing while the high
// bit is set. (Multiplication, not `<<`, so a high byte at shift ≥ 24 can't go negative in JS's
// 32-bit bitwise domain.)
function readId(bytes: Uint8Array, pos: number): { value: number; next: number } | undefined {
	let value = 0;
	for (let i = 0; i < 4; i++) {
		if (pos >= bytes.length) return undefined;
		const b = bytes[pos] as number;
		pos += 1;
		value += b * 2 ** (8 * i);
		if ((b & 0x80) === 0) return { value, next: pos };
	}
	return { value, next: pos };
}

// Read the payload length: a 7-bits-per-byte varint, continuing while the high bit is set.
function readLen(bytes: Uint8Array, pos: number): { value: number; next: number } | undefined {
	let value = 0;
	for (let i = 0; i < 4; i++) {
		if (pos >= bytes.length) return undefined;
		const b = bytes[pos] as number;
		pos += 1;
		value += (b & 0x7f) * 2 ** (7 * i);
		if ((b & 0x80) === 0) return { value, next: pos };
	}
	return { value, next: pos };
}

/** Walk every record in a BIFF12 part. Stops cleanly at end-of-buffer or a truncated header. */
export function* readRecords(bytes: Uint8Array): Generator<BiffRecord> {
	let pos = 0;
	while (pos < bytes.length) {
		const id = readId(bytes, pos);
		if (id === undefined) return;
		const len = readLen(bytes, id.next);
		if (len === undefined) return;
		const start = len.next;
		// A declared length that overruns the buffer is a truncated final record — take what's left.
		const end = Math.min(start + len.value, bytes.length);
		yield { id: id.value, data: bytes.subarray(start, end) };
		pos = end;
	}
}

/**
 * A cursor for reading typed fields out of a single record's payload. Every read is bounds-checked
 * against the record: past the end, integers read 0 and strings read empty, so a short or malformed
 * record degrades instead of throwing.
 */
export class RecordData {
	readonly #bytes: Uint8Array;
	readonly #view: DataView;
	#pos = 0;

	constructor(data: Uint8Array) {
		this.#bytes = data;
		this.#view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	}

	get remaining(): number {
		return this.#bytes.length - this.#pos;
	}

	skip(n: number): void {
		this.#pos += n;
	}

	u8(): number {
		if (this.remaining < 1) return 0;
		return this.#view.getUint8(this.#pos++);
	}

	u16(): number {
		if (this.remaining < 2) return 0;
		const v = this.#view.getUint16(this.#pos, true);
		this.#pos += 2;
		return v;
	}

	u32(): number {
		if (this.remaining < 4) return 0;
		const v = this.#view.getUint32(this.#pos, true);
		this.#pos += 4;
		return v;
	}

	f64(): number {
		if (this.remaining < 8) return 0;
		const v = this.#view.getFloat64(this.#pos, true);
		this.#pos += 8;
		return v;
	}

	/** Decode an RK number: a packed int-or-truncated-double with an optional ÷100 flag. */
	rk(): number {
		if (this.remaining < 4) return 0;
		const iv = this.#view.getInt32(this.#pos, true);
		this.#pos += 4;
		let v: number;
		if (iv & 0x02) {
			v = iv >> 2; // integer: arithmetic shift keeps the sign
		} else {
			// The 4 RK bytes (low 2 bits cleared) are the HIGH 4 bytes of an 8-byte double.
			const buf = new ArrayBuffer(8);
			const dv = new DataView(buf);
			dv.setInt32(4, iv & ~0x03, true);
			v = dv.getFloat64(0, true);
		}
		return iv & 0x01 ? v / 100 : v;
	}

	/**
	 * Read an XL(Nullable)WideString: a uint32 count of UTF-16 code units, then that many code units
	 * (2 bytes each). `0xFFFFFFFF` is the null-string sentinel → `undefined`.
	 */
	wideString(): string | undefined {
		if (this.remaining < 4) return undefined;
		const count = this.#view.getUint32(this.#pos, true);
		this.#pos += 4;
		if (count === 0xffffffff) return undefined;
		const byteLen = Math.min(count * 2, this.remaining);
		const s = utf16.decode(this.#bytes.subarray(this.#pos, this.#pos + byteLen));
		this.#pos += byteLen;
		return s;
	}
}
