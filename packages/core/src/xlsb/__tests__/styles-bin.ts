import { CELLXFS, CELLXFS_END, FMT, XF } from "../records";

// BIFF12 builders for styles.bin pins: byte-combined record ids, 7-bit varint lengths, little-endian
// integers and XLWideString.

/**
 * A byte-combined record id as its one or two wire bytes.
 *
 * @example recordId(0x2c) // [0x2c]
 * @example recordId(0x04e9) // [0xe9, 0x04]
 */
export function recordId(id: number): number[] {
	return id < 0x80 ? [id] : [id & 0xff, id >> 8];
}

/** A record length as a 7-bit varint. @example recordLength(200) // [0xc8, 0x01] */
export function recordLength(length: number): number[] {
	const out: number[] = [];
	let rest = length;
	do {
		let byte = rest & 0x7f;
		rest >>>= 7;
		if (rest) byte |= 0x80;
		out.push(byte);
	} while (rest);
	return out;
}

/** A little-endian u16. @example u16(0x0102) // [0x02, 0x01] */
export function u16(value: number): number[] {
	return [value & 0xff, (value >> 8) & 0xff];
}

/** A little-endian u32. @example u32(1) // [1, 0, 0, 0] */
export function u32(value: number): number[] {
	return [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff, (value >>> 24) & 0xff];
}

/** An XLWideString: a u32 character count, then UTF-16LE code units. @example wideString("A") // [1, 0, 0, 0, 0x41, 0] */
export function wideString(text: string): number[] {
	const out = u32(text.length);
	for (let i = 0; i < text.length; i++) out.push(...u16(text.charCodeAt(i)));
	return out;
}

/** One record: id, length, payload. @example record(0x2c, [1]) // [0x2c, 1, 1] */
export function record(id: number, payload: readonly number[]): number[] {
	return [...recordId(id), ...recordLength(payload.length), ...payload];
}

/**
 * A styles.bin with BrtFmt records for `formats` ([numFmtId, code] pairs) and one BrtXF per entry of
 * `xfFormats` (each xf's numFmtId), inside BrtBeginCellXFs/BrtEndCellXFs.
 *
 * @example stylesBin([[164, "0.0"]], [0, 164]) // two formats-in-use: General and id 164
 */
export function stylesBin(
	formats: readonly (readonly [number, string])[],
	xfFormats: readonly number[],
): Uint8Array {
	const bytes: number[] = [];
	for (const [id, code] of formats) bytes.push(...record(FMT, [...u16(id), ...wideString(code)]));
	bytes.push(...record(CELLXFS, u32(xfFormats.length)));
	for (const id of xfFormats) bytes.push(...record(XF, [...u16(0), ...u16(id), ...u32(0)]));
	bytes.push(...record(CELLXFS_END, []));
	return Uint8Array.from(bytes);
}
