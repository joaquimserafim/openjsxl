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

/** One populated cell surfaced by a {@link RangeView}: its 1-based position and evaluated value. */
export interface RangeEntry {
	readonly col: number;
	readonly row: number;
	readonly value: EvalValue;
}

/**
 * A lazy view over the cells of a reference. It never materializes the rectangle: `width`/`height`
 * are computed from the bounds (so `A:A` reports 1 048 576 rows for free), and iteration visits only
 * the USED cells the evaluator supplies, resolving each on demand (a formula cell is evaluated when
 * first pulled). Aggregators (F8.3) consume `entries()`/`values()`; `COUNTBLANK` uses the extent
 * arithmetic `cellCount - populatedCount()`.
 */
export class RangeView {
	constructor(
		/** Resolved sheet name the range lives on. */
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

	/** Populated (non-blank) cells, evaluated on demand, in the model's iteration order. */
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
}
