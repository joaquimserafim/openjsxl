import { MAX_COL } from "../ooxml/a1";

// Internal column-style compiler shared by row assembly and the worksheet style scan. A declaration
// is applied once; cell lookup is direct and never scans the original ranges.

/** A column-level default cell style declared by `<col min max style>`. */
export interface ColumnStyle {
	readonly min: number;
	readonly max: number;
	readonly style: number;
}

export interface ColumnStyleResolver {
	/** Input range 2..4/style 7; output is the number of newly assigned grid columns. */
	add(range: ColumnStyle): number;
	/** Input column 3 after that range; output style 7. Out-of-grid columns return `undefined`. */
	styleFor(column: number | undefined): number | undefined;
}

/** Input `{ min: "2", max: "4", style: "7" }`; output `{ min: 2, max: 4, style: 7 }`. */
export function columnStyleFromToken(
	attrs: Readonly<Record<string, string>>,
): ColumnStyle | undefined {
	if (attrs.style === undefined) return undefined;
	const style = Number(attrs.style);
	const min = Number(attrs.min);
	const max = Number(attrs.max);
	if (!Number.isInteger(style) || !Number.isInteger(min) || !Number.isInteger(max)) {
		return undefined;
	}
	return { min, max, style };
}

// Successor-with-deletion: once a grid column has its first style, link it to the next unassigned
// column. Path compression means all declarations together visit each grid column only once, apart
// from near-constant successor lookups.
/** Input links `1 → 2 → 3` and start 1; output 3 and compress both earlier links to 3. */
export function nextUnassigned(links: Uint32Array, start: number): number {
	let root = start;
	while ((links[root] as number) !== root) root = links[root] as number;
	let column = start;
	while ((links[column] as number) !== column) {
		const parent = links[column] as number;
		links[column] = root;
		column = parent;
	}
	return root;
}

/**
 * Compile first-covering column defaults into Excel's fixed grid.
 * Example: add 1..2/style 7, then 2..3/style 8; output styles are `[7, 7, 8]` for columns 1..3.
 */
export function createColumnStyleResolver(): ColumnStyleResolver {
	let styles: Array<number | undefined> | undefined;
	let links: Uint32Array | undefined;

	return {
		add(range) {
			const { min, max, style } = range;
			if (!Number.isInteger(min) || !Number.isInteger(max) || !Number.isInteger(style)) {
				return 0;
			}
			const first = Math.max(1, min);
			const last = Math.min(MAX_COL, max);
			if (first > last) return 0;

			let values = styles;
			let successors = links;
			if (values === undefined || successors === undefined) {
				values = new Array<number | undefined>(MAX_COL + 1);
				successors = new Uint32Array(MAX_COL + 2);
				for (let column = 0; column < successors.length; column++) {
					successors[column] = column;
				}
				styles = values;
				links = successors;
			}

			let assigned = 0;
			let column = nextUnassigned(successors, first);
			while (column <= last) {
				values[column] = style;
				successors[column] = nextUnassigned(successors, column + 1);
				column = successors[column] as number;
				assigned++;
			}
			return assigned;
		},

		styleFor(column) {
			if (
				styles === undefined ||
				column === undefined ||
				!Number.isInteger(column) ||
				column < 1 ||
				column > MAX_COL
			) {
				return undefined;
			}
			return styles[column];
		},
	};
}
