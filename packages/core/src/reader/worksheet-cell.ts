import type { ColumnStyleResolver } from "./worksheet-column-styles";

// Shared cell address/style rules used by row assembly and the worksheet metadata scans.
// Keeping them here prevents cell(), numberFormat(), and formula() from interpreting the same
// worksheet position differently.

/** Input `C99`; output `3`. Malformed or unsafe refs return `undefined` without exceptions. */
export function safeColumn(ref: string): number | undefined {
	const match = /^([A-Za-z]+)[1-9][0-9]*$/.exec(ref);
	if (match === null) return undefined;
	const letters = match[1] as string;
	let column = 0;
	for (let i = 0; i < letters.length; i++) {
		const code = letters.charCodeAt(i);
		const value = code >= 97 ? code - 96 : code - 64;
		if (column > Math.floor((Number.MAX_SAFE_INTEGER - value) / 26)) return undefined;
		column = column * 26 + value;
	}
	return column;
}

/** Input `{ s: "5", customFormat: "1" }`; output style `5`. Without the flag: `undefined`. */
export function rowDefaultStyleFromToken(
	attrs: Readonly<Record<string, string>>,
): number | undefined {
	if (attrs.customFormat !== "1" && attrs.customFormat !== "true") return undefined;
	if (attrs.s === undefined) return undefined;
	const style = Number(attrs.s);
	return Number.isInteger(style) ? style : undefined;
}

/** Input own style `3`, row `5`, column `7`; output `3` (cell wins). */
export function effectiveStyle(
	ownStyle: string | undefined,
	rowDefault: number | undefined,
	column: number | undefined,
	columns: ColumnStyleResolver,
): number | undefined {
	if (ownStyle !== undefined) {
		const style = Number(ownStyle);
		return Number.isInteger(style) ? style : undefined;
	}
	if (rowDefault !== undefined) return rowDefault;
	return columns.styleFor(column);
}
