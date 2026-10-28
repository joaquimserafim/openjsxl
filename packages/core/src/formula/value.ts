// The runtime value domain of the evaluator (F8.2). Distinct from the parse-time AST: an `EvalValue`
// is a *result*, not syntax. Errors are ordinary values here — they propagate rather than throw
// (decision 4) — so a `#DIV/0!` in one cell never aborts an unrelated cell.

/**
 * A cell error carried as a VALUE. The eight members are the `ST_CellErrorType` set; `#CYCLE!` is an
 * evaluation-only addition for circular references (decision 4 — HyperFormula precedent, never
 * Excel's silent 0). Evaluation is read-only, so a `#CYCLE!` result is never written back to a cell.
 */
export type ErrorCode =
	| "#DIV/0!"
	| "#VALUE!"
	| "#REF!"
	| "#NAME?"
	| "#N/A"
	| "#NUM!"
	| "#NULL!"
	| "#GETTING_DATA"
	| "#CYCLE!";

export interface FormulaErrorValue {
	readonly kind: "error";
	readonly code: ErrorCode;
}

// Interned instances — every error value of a given code is the same frozen object, so equality is
// cheap and no allocation happens on the hot propagation path.
const ERROR_VALUES: Record<ErrorCode, FormulaErrorValue> = {
	"#DIV/0!": Object.freeze({ kind: "error", code: "#DIV/0!" }),
	"#VALUE!": Object.freeze({ kind: "error", code: "#VALUE!" }),
	"#REF!": Object.freeze({ kind: "error", code: "#REF!" }),
	"#NAME?": Object.freeze({ kind: "error", code: "#NAME?" }),
	"#N/A": Object.freeze({ kind: "error", code: "#N/A" }),
	"#NUM!": Object.freeze({ kind: "error", code: "#NUM!" }),
	"#NULL!": Object.freeze({ kind: "error", code: "#NULL!" }),
	"#GETTING_DATA": Object.freeze({ kind: "error", code: "#GETTING_DATA" }),
	"#CYCLE!": Object.freeze({ kind: "error", code: "#CYCLE!" }),
};

/** The interned error value for a code. */
export function errorValue(code: ErrorCode): FormulaErrorValue {
	return ERROR_VALUES[code];
}

/**
 * Map an arbitrary error string (from a reader error cell or a parsed error literal) to an error
 * value. An unrecognized string degrades to `#VALUE!` rather than fabricating a new error code.
 */
export function errorFromString(code: string): FormulaErrorValue {
	switch (code) {
		case "#DIV/0!":
		case "#VALUE!":
		case "#REF!":
		case "#NAME?":
		case "#N/A":
		case "#NUM!":
		case "#NULL!":
		case "#GETTING_DATA":
		case "#CYCLE!":
			return ERROR_VALUES[code];
		default:
			return ERROR_VALUES["#VALUE!"];
	}
}

/** A single (non-range) evaluation result. `null` is an empty cell. */
export type ScalarValue = number | string | boolean | null | FormulaErrorValue;

/**
 * A scalar-or-range evaluation result. `null` is an empty cell (coerces to 0 / `""` per decision 6);
 * a {@link RangeView} is a lazy window over a reference's USED cells — never a materialized array.
 */
export type EvalValue = ScalarValue | RangeView;

export function isErrorValue(v: unknown): v is FormulaErrorValue {
	return (
		typeof v === "object" &&
		v !== null &&
		v instanceof Object &&
		"kind" in v &&
		v.kind === "error"
	);
}

export function isRangeView(v: unknown): v is RangeView {
	return v instanceof RangeView;
}

/**
 * The value with an empty cell read as the 0 Excel shows in its place: what a formula cell holds when
 * its result is an empty cell (F11.62). VLOOKUP and HLOOKUP also apply it to an empty target.
 *
 * @example blankToZero(null) // 0
 * @example blankToZero("") // "" (an empty string is not an empty cell)
 */
export function blankToZero(v: EvalValue): EvalValue {
	return v === null ? 0 : v;
}

/**
 * The value with an empty cell read as an empty string: what IFERROR and IFNA return for an empty
 * value or fallback, which Microsoft documents as "treats it as an empty string value" (F11.62). A
 * range over one empty cell is such a value; any other value, a wider range included, is kept as is.
 *
 * @example blankToText(null) // ""
 * @example blankToText(new RangeView("S", 1, 1, 1, 1, () => [], () => null)) // "" (one empty cell)
 * @example blankToText(0) // 0
 */
export function blankToText(v: EvalValue): EvalValue {
	if (v === null) return "";
	if (isRangeView(v) && v.single() === null) return "";
	return v;
}

/**
 * The most characters a text value holds in Excel: 32,767, its cell limit. `&` and the built-ins that
 * build text give `#VALUE!` for a longer result, as Excel does (F11.66); `=A1` still returns a longer
 * text that a reader stored, as it is.
 */
export const MAX_TEXT = 32767;

/**
 * The text, or `#VALUE!` when it is longer than {@link MAX_TEXT} characters (F11.66).
 *
 * @example capText("abc") // "abc"
 * @example capText("a".repeat(32768)) // #VALUE!
 */
export function capText(text: string): string | FormulaErrorValue {
	return text.length > MAX_TEXT ? errorValue("#VALUE!") : text;
}

/** One populated cell surfaced by a {@link RangeView}: its 1-based position and evaluated value. */
export interface RangeEntry {
	readonly col: number;
	readonly row: number;
	readonly value: EvalValue;
}

// Keep coordinate-only traversal internal: conditional aggregates must check criteria before
// resolving a candidate value. Public entries() retains its existing evaluation timing.
let rangePositions: (range: RangeView) => Iterable<readonly [number, number]>;

/** @internal Populated coordinates without evaluating cells; not part of the public entry point. */
export function populatedRangePositions(range: RangeView): Iterable<readonly [number, number]> {
	return rangePositions(range);
}

/**
 * A lazy view over the cells of a reference. It never materializes the rectangle: `width`/`height`
 * are computed from the bounds (so `A:A` reports 1 048 576 rows for free), and iteration visits only
 * the USED cells the evaluator supplies, resolving each on demand (a formula cell is evaluated when
 * first pulled). Aggregators (F8.3) consume `entries()`/`values()`; `COUNTBLANK` uses the extent
 * arithmetic `cellCount - populatedCount()`.
 */
export class RangeView {
	static {
		// A module-private accessor keeps populated coordinates out of the public class API.
		// Use the same receiver as entries(), including for caller-provided Proxy wrappers.
		rangePositions = (range) => range.populated();
	}

	constructor(
		/** Resolved sheet name the range lives on; `""` for an array constant (F11.64). */
		readonly sheet: string,
		/** 1-based inclusive bounds. */
		readonly startCol: number,
		readonly startRow: number,
		readonly endCol: number,
		readonly endRow: number,
		// Coordinates of the populated cells within the bounds (lazy — the evaluator streams them
		// from its sparse model). Each is a [col, row] pair.
		private readonly populated: () => Iterable<readonly [number, number]>,
		// Resolve one cell to its value (evaluating a formula cell through the shared walker).
		private readonly resolve: (col: number, row: number) => EvalValue,
	) {}

	get width(): number {
		return this.endCol - this.startCol + 1;
	}

	get height(): number {
		return this.endRow - this.startRow + 1;
	}

	/** Total addressable cells in the rectangle (blank included) — extent arithmetic, never iterated. */
	get cellCount(): number {
		return this.width * this.height;
	}

	/**
	 * Populated (non-blank) cells, evaluated on demand, in the order the view's source gives them: row
	 * by row and left to right for a range the evaluator builds (F11.67) and for an array constant.
	 */
	*entries(): IterableIterator<RangeEntry> {
		for (const [col, row] of this.populated()) {
			yield { col, row, value: this.resolve(col, row) };
		}
	}

	/** Just the values of the populated cells. */
	*values(): IterableIterator<EvalValue> {
		for (const entry of this.entries()) yield entry.value;
	}

	/** How many cells are populated (non-blank). Iterates the model's used set, not the rectangle. */
	populatedCount(): number {
		let n = 0;
		for (const _ of this.populated()) n++;
		return n;
	}

	/**
	 * The value of a single-cell range (`A1:A1`, or an `@`/implicit-intersection reduction to one
	 * cell), or `undefined` when the range covers more than one cell. An empty single cell is `null`.
	 */
	single(): EvalValue | undefined {
		if (this.width !== 1 || this.height !== 1) return undefined;
		return this.resolve(this.startCol, this.startRow);
	}

	/** The value of the top-left cell — how a cell holding a range formula displays (decision 5). */
	topLeft(): EvalValue {
		return this.resolve(this.startCol, this.startRow);
	}

	/**
	 * The value at a 0-based offset within the view (`[rowOffset, colOffset]` from the top-left), or
	 * `null` (blank) when the offset falls outside the rectangle. The positional accessor the function
	 * library uses to align a criteria range with a sum range and to index INDEX/VLOOKUP/MATCH results.
	 */
	cellAt(rowOffset: number, colOffset: number): EvalValue {
		if (rowOffset < 0 || colOffset < 0 || rowOffset >= this.height || colOffset >= this.width) {
			return null;
		}
		return this.resolve(this.startCol + colOffset, this.startRow + rowOffset);
	}
}

/**
 * Every [col, row] position of a `width` × `height` grid from (1, 1), row by row.
 *
 * @example [...gridPositions(2, 2)] // [[1, 1], [2, 1], [1, 2], [2, 2]]
 * @example [...gridPositions(0, 3)] // []
 */
export function* gridPositions(
	width: number,
	height: number,
): Generator<readonly [number, number]> {
	for (let row = 1; row <= height; row++) {
		for (let col = 1; col <= width; col++) yield [col, row];
	}
}

/**
 * An array constant (`{1,2;3,4}`) as a range over its own elements, so a function argument sees every
 * element (F11.64). Every position is populated, and `sheet` is empty because it names no sheet.
 * Where one value is wanted, it stands for its first element: the value Excel shows for an array
 * result (decision 5).
 *
 * @example new ArrayView([[1, 2], [3, 4]]).cellAt(1, 0) // 3
 */
export class ArrayView extends RangeView {
	constructor(rows: readonly (readonly EvalValue[])[]) {
		const width = rows[0]?.length ?? 0;
		super(
			"",
			1,
			1,
			width,
			rows.length,
			() => gridPositions(width, rows.length),
			(col, row) => rows[row - 1]?.[col - 1] ?? null,
		);
	}
}

/**
 * The value with an array constant read as its first element, the value a cell holding an array
 * result shows (decision 5, F11.64). Any other value, a reference range included, is kept as is.
 *
 * @example anchorValue(new ArrayView([[1, 2]])) // 1
 * @example anchorValue(5) // 5
 */
export function anchorValue(v: EvalValue): EvalValue {
	return v instanceof ArrayView ? v.topLeft() : v;
}

/**
 * A new array constant of the same shape, with each element replaced by `fn(element, row, col)`, the
 * position given as 0-based offsets.
 *
 * @example mapArray(new ArrayView([[1, 2]]), (v) => (v === 1 ? 9 : v)) // an ArrayView over 9, 2
 */
export function mapArray(
	view: ArrayView,
	fn: (value: EvalValue, row: number, col: number) => EvalValue,
): ArrayView {
	const rows: EvalValue[][] = [];
	for (let row = 0; row < view.height; row++) {
		const values: EvalValue[] = [];
		for (let col = 0; col < view.width; col++) values.push(fn(view.cellAt(row, col), row, col));
		rows.push(values);
	}
	return new ArrayView(rows);
}

/**
 * One row, one column, or one element of an array constant, as a new array constant. `row` and `col`
 * are 0-based offsets inside the array; `undefined` keeps every row or every column.
 *
 * @example sliceArray(new ArrayView([[1, 2], [3, 4]]), undefined, 0) // an ArrayView over 1, 3
 * @example sliceArray(new ArrayView([[1, 2], [3, 4]]), 1, undefined) // an ArrayView over 3, 4
 */
export function sliceArray(
	view: ArrayView,
	row: number | undefined,
	col: number | undefined,
): ArrayView {
	const rows: EvalValue[][] = [];
	for (let r = row ?? 0; r <= (row ?? view.height - 1); r++) {
		const values: EvalValue[] = [];
		for (let c = col ?? 0; c <= (col ?? view.width - 1); c++) values.push(view.cellAt(r, c));
		rows.push(values);
	}
	return new ArrayView(rows);
}
