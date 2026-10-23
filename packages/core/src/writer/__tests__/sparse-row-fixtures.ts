import { createMediaRegistry } from "../images";
import { createTableContext, worksheetXml } from "../sheet";
import { createStyleRegistry } from "../styles";
import type { CellInput, SheetInput } from "../types";

// F11.54 fixtures: holey rows with far-right cells, and a buffered render of one sheet.

/**
 * `count` holey rows; row `r` holds only the number `r`, at 1-based `column`.
 *
 * @example farRightRows(2, 3) // [[<2 holes>, 0], [<2 holes>, 1]]
 */
export function farRightRows(count: number, column: number): CellInput[][] {
	return Array.from({ length: count }, (_, r) => {
		const row: CellInput[] = [];
		row[column - 1] = r;
		return row;
	});
}

/**
 * The worksheet XML the buffered writer renders for `sheet`, with fresh style, media and table
 * registries.
 *
 * @example sheetXml({ name: "S", rows: [[1]] }) // '<?xml …><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
 */
export function sheetXml(sheet: SheetInput): string {
	return worksheetXml(
		sheet,
		sheet.rows,
		0,
		false,
		createStyleRegistry(),
		createMediaRegistry(),
		createTableContext(),
		sheet.name,
	).xml;
}

/** How often a row was probed for one own index, and how often its keys were listed. */
export interface RowReads {
	probes: number;
	keyListings: number;
}

/**
 * A Proxy over `row` that counts own-index probes (`Object.hasOwn`) and key listings
 * (`Object.getOwnPropertyNames`) in `reads`. `reverseKeys` lists the keys in reverse order, as a
 * Proxy may.
 *
 * @example countingRow([1, 2], reads) // reads.probes grows by one per Object.hasOwn(row, i)
 */
export function countingRow(row: CellInput[], reads: RowReads, reverseKeys = false): CellInput[] {
	return new Proxy(row, {
		getOwnPropertyDescriptor(target, key) {
			if (typeof key === "string" && /^\d+$/.test(key)) reads.probes++;
			return Reflect.getOwnPropertyDescriptor(target, key);
		},
		ownKeys(target) {
			reads.keyListings++;
			const keys = Reflect.ownKeys(target);
			return reverseKeys ? keys.reverse() : keys;
		},
	});
}
