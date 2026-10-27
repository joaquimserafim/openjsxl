import { BLANK, NUM, ROW, SHORT_BLANK, SHORT_NUM } from "../records";
import { record, u32 } from "./styles-bin";

// BIFF12 builders for sheetN.bin pins (F11.59): rows, RK numbers and blanks, and large sheets of
// one repeated cell record.

/**
 * A BrtRow record for the 0-based `row`.
 *
 * @example rowRecord(2) // [0x00, 4, 2, 0, 0, 0]
 */
export function rowRecord(row: number): number[] {
	return record(ROW, u32(row));
}

/**
 * A BrtShortRk record (the next column) holding the integer `value`, with style `style`.
 *
 * @example shortRk(7, 2) // [0x0d, 8, 2, 0, 0, 0, 0x1e, 0, 0, 0]
 */
export function shortRk(value: number, style = 0): number[] {
	return record(SHORT_NUM, [...u32(style), ...u32((value << 2) | 2)]);
}

/**
 * A BrtShortBlank record (the next column) with style `style`. With no style it has no payload, the
 * 2-byte form whose style reads as 0.
 *
 * @example shortBlank() // [0x0c, 0]
 * @example shortBlank(3) // [0x0c, 4, 3, 0, 0, 0]
 */
export function shortBlank(style?: number): number[] {
	return style === undefined ? record(SHORT_BLANK, []) : record(SHORT_BLANK, u32(style));
}

/**
 * A BrtCellRk record at the 0-based `col` holding the integer `value`, with style `style`.
 *
 * @example fullRk(1, 7, 2) // [0x02, 12, 1, 0, 0, 0, 2, 0, 0, 0, 0x1e, 0, 0, 0]
 */
export function fullRk(col: number, value: number, style = 0): number[] {
	return record(NUM, [...u32(col), ...u32(style), ...u32((value << 2) | 2)]);
}

/**
 * A BrtCellBlank record at the 0-based `col` with style `style`.
 *
 * @example fullBlank(1, 3) // [0x01, 8, 1, 0, 0, 0, 3, 0, 0, 0]
 */
export function fullBlank(col: number, style = 0): number[] {
	return record(BLANK, [...u32(col), ...u32(style)]);
}

/**
 * A sheet part of `rows` rows, each a BrtRow then `cell` repeated `cols` times. It is built by
 * copying the cell's bytes, so a sheet of millions of records builds quickly.
 *
 * @example repeatedCellsBin(2, 3, shortBlank()) // two rows (6-byte BrtRow) of three 2-byte blanks: 24 bytes
 */
export function repeatedCellsBin(rows: number, cols: number, cell: readonly number[]): Uint8Array {
	const rowBytes = rowRecord(0).length;
	const out = new Uint8Array(rows * (rowBytes + cols * cell.length));
	const cellBytes = Uint8Array.from(cell);
	let at = 0;
	for (let r = 0; r < rows; r++) {
		out.set(rowRecord(r), at);
		at += rowBytes;
		for (let c = 0; c < cols; c++) {
			out.set(cellBytes, at);
			at += cellBytes.length;
		}
	}
	return out;
}
