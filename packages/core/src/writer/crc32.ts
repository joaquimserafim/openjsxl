// CRC-32 (IEEE 802.3) — the per-entry checksum every ZIP local + central header carries.
//
// The reader (zip/central-directory.ts) never verifies this: it trusts the archive and slices
// payloads by the recorded sizes. A writer, though, MUST get it right — Excel, `unzip`, and any
// strict reader reject an entry whose stored CRC doesn't match its bytes, so a wrong value here
// is the difference between "opens" and "corrupt file".
//
// Table-driven: the 256-entry lookup is built once (reflected polynomial 0xEDB88320), then the
// checksum is a single branch-free pass over the bytes. This is the exact variant ZIP, gzip, and
// PNG share — initial value all-ones, final one's-complement — computed over the *uncompressed*
// bytes regardless of whether the entry is later stored or deflated.

const CRC_TABLE = (() => {
	const table = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		table[n] = c >>> 0;
	}
	return table;
})();

/** The all-ones seed a running CRC-32 starts from (before {@link crc32Update}). */
export const CRC32_INIT = 0xffffffff;

/**
 * Fold `bytes` into a running CRC-32 accumulator. The accumulator is the internal (pre-final-XOR)
 * state, so streaming callers thread it across chunks and call {@link crc32Final} once at the end —
 * exactly what a ZIP entry written with a data descriptor needs, since its bytes arrive in pieces.
 */
export function crc32Update(crc: number, bytes: Uint8Array): number {
	for (let i = 0; i < bytes.length; i++) {
		// `& 0xff` keeps the low byte; `>>> 8` shifts the running value; final `>>> 0` forces the
		// result back to an unsigned 32-bit int (^ in JS yields a signed number).
		crc = ((CRC_TABLE[(crc ^ (bytes[i] as number)) & 0xff] as number) ^ (crc >>> 8)) >>> 0;
	}
	return crc;
}

/** Apply the final one's-complement to a running accumulator, giving the header-ready checksum. */
export function crc32Final(crc: number): number {
	return (crc ^ 0xffffffff) >>> 0;
}

/** CRC-32 of `bytes` as an unsigned 32-bit integer, ready to write little-endian into a header. */
export function crc32(bytes: Uint8Array): number {
	return crc32Final(crc32Update(CRC32_INIT, bytes));
}
