import { columnToIndex, indexToColumn, MAX_COL, MAX_ROW } from "../ooxml/a1";
import { dateToSerial } from "../ooxml/dates";
import type { DefinedName } from "../ooxml/workbook";
import type { Workbook } from "../reader/workbook";
import type { Cell, Worksheet } from "../types";
import type { ArrayElement, ArrayLiteral, FormulaAst } from "./ast";
import { applyBinary, applyPostfix, applyUnary, scalarize } from "./coerce";
import { FormulaError } from "./errors";
import {
	buildRegistry,
	type EvalContext,
	type FunctionRegistry,
	lookupFunction,
} from "./functions";
import { parseFormula } from "./parser";
import {
	ArrayView,
	anchorValue,
	blankToZero,
	type EvalValue,
	errorFromString,
	errorValue,
	isRangeView,
	RangeView,
	type ScalarValue,
} from "./value";

// The evaluator core (F8.2). A one-shot, pull-based, memoizing tri-color walker over the shared
// Workbook surface. Two properties are load-bearing:
//   • Deep dependency CHAINS never grow the JS stack. Each cell's formula is a generator that YIELDS
//     the cells it needs; an explicit frame stack (runFormula) resolves them, so a 1M-row chain of
//     `=A1+1` costs 1M heap frames and O(1) native stack.
//   • CYCLES are values, not aborts. Tri-color marks (white = unseen, grey = in progress, black =
//     done) give topological order implicitly and detect a cycle for free — a grey hit yields the
//     dedicated `#CYCLE!` value for that back-edge, so the cycle's cells become `#CYCLE!` while every
//     unrelated cell evaluates normally.
// Functions (F8.3) are dispatched through the registry; F8.2 proves the machinery with test specs.

/** Options for {@link evaluateWorkbook} / {@link evaluateCell}. */
export interface EvaluateOptions {
	/**
	 * Extra functions, keyed by name (case-insensitive), merged over the built-ins. Each value is a
	 * {@link import("./functions").FunctionSpec}. Validated as a plain object with numeric arg bounds
	 * and an `evaluate` function. A call Excel stores as `_xlfn.NAME` or `_xlfn._xlws.NAME` reaches the
	 * function registered as `NAME`, unless one is registered under the prefixed name itself. An array
	 * constant argument (`{1,2;3,4}`) arrives as a `RangeView` of its elements whose `sheet` is `""`.
	 */
	readonly functions?: Record<string, unknown>;
	/** Injected clock for volatile NOW/TODAY (decision 3). Without it, a volatile call is rejected. */
	readonly now?: () => Date;
	/** Injected RNG for volatile RAND/RANDBETWEEN. Without it, a volatile call is rejected. */
	readonly random?: () => number;
	/** Fuel budget: the maximum number of formula-cell evaluations before a typed abort. */
	readonly maxCellVisits?: number;
}

/** One formula cell's evaluated result. */
export interface CellResult {
	readonly sheet: string;
	readonly ref: string;
	readonly value: ScalarValue;
}

/** The result of evaluating a whole workbook — every formula cell, plus a keyed lookup. */
export interface WorkbookEvalResult {
	readonly cells: readonly CellResult[];
	/**
	 * The evaluated value at `sheet!ref`, or `undefined` when that cell holds no formula. The sheet and
	 * ref match as `cells` spells them (`"Sheet1"`, `"B2"`); `evaluateCell` also takes another case.
	 */
	get(sheet: string, ref: string): ScalarValue | undefined;
}

// A cell in the snapshot model: a literal value, or a formula's stored text (parsed on first use).
type ModelCell =
	| { readonly kind: "value"; readonly value: EvalValue }
	| { readonly kind: "formula"; readonly text: string };

interface SheetModel {
	readonly worksheet: Worksheet;
	readonly index: number;
	// row → (col → cell). Only POPULATED cells; a miss is an empty cell (or a formula-only cell,
	// recovered via worksheet.formula on demand).
	readonly cells: Map<number, Map<number, ModelCell>>;
}

// Tri-color state: absent = white, "grey" = on the current evaluation path, {value} = black (memoized).
type CellState = "grey" | { readonly value: EvalValue };

// What a formula name resolves to, pre-classified (decision 5: constants + simple refs only).
type ResolvedName =
	| { readonly kind: "const"; readonly value: ScalarValue }
	| { readonly kind: "ref"; readonly ast: FormulaAst; readonly sheet: string }
	| { readonly kind: "unsupported" };

const A1_REF = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COL_ONLY = /^\$?([A-Za-z]{1,3})$/;
const ROW_ONLY = /^\$?(\d{1,7})$/;

// A single-cell dependency the generator asks the driver to resolve.
interface CellRequest {
	readonly sheet: string;
	readonly col: number;
	readonly row: number;
}

// A range endpoint's partial address (`sheet?: string | undefined` so an unqualified endpoint is
// representable under exactOptionalPropertyTypes).
interface Endpoint {
	readonly col?: number;
	readonly row?: number;
	readonly sheet?: string | undefined;
}

const DEFAULT_MAX_VISITS = 5_000_000;
// Native-recursion guard. `depth` is threaded through the whole synchronous recursion — AST descent
// within a cell (evalGen) AND cross-cell nesting through lazy-argument thunks / range iteration — so
// a single bound covers both. Crucially it does NOT count the iterative frame stack: a simple
// reference chain resets to the driver's base depth per frame, so a 1M-row `=A1+1` chain stays at
// depth ~2 while a 4096-term `1+1+…` chain (left-assoc, which the parser builds iteratively and so
// escapes ITS depth cap) trips this and becomes a typed error instead of a native stack overflow.
// Kept conservative so the guard fires well before V8's own stack limit on any host (the limit
// varies by platform); real formulas nest far shallower than this.
const MAX_NATIVE = 256;

function keyOf(sheet: string, col: number, row: number): string {
	return `${sheet}\u0000${col}\u0000${row}`;
}

function parseA1(ref: string): { col: number; row: number } | undefined {
	const m = A1_REF.exec(ref);
	if (m === null) return undefined;
	return { col: columnToIndex(m[1] as string), row: Number.parseInt(m[2] as string, 10) };
}

// Convert a reader Cell into an EvalValue. A date becomes its serial (the number Excel arithmetic
// sees) in the WORKBOOK's date system (F11.10): a 1904-epoch file's 1904-01-01 is serial 0, not 1462.
function cellToValue(cell: Cell, date1904: boolean): EvalValue {
	switch (cell.type) {
		case "empty":
			return null;
		case "number":
			return cell.value;
		case "string":
			return cell.value;
		case "boolean":
			return cell.value;
		case "date":
			return dateToSerial(cell.value, date1904);
		case "error":
			return errorFromString(cell.value);
	}
}

/**
 * Whether a cell the reader returns belongs in the evaluator's model: it holds a formula or a value. A
 * cell the reader reads as empty, with no formula it can return, is empty to every formula, as
 * `Worksheet.cell()` reports it. That covers a cell that carries only a style, so formatting a cell
 * never changes a result, as in Excel (F11.63), and a value the reader degrades to empty.
 *
 * @example holdsValue({ ref: "A1", type: "number", value: 3 }, undefined) // true
 * @example holdsValue({ ref: "A1", type: "empty", value: null }, "B1*2") // true (a formula cell)
 * @example holdsValue({ ref: "A1", type: "empty", value: null }, undefined) // false (a styled blank)
 */
export function holdsValue(cell: Cell, formula: string | undefined): boolean {
	return formula !== undefined || cell.type !== "empty";
}

/**
 * The sheet a formula's qualifier names (`data!A1`, `'DATA'!Rate`), matched ignoring case as Excel
 * does, with the fold the reader uses for duplicate names (`toLowerCase`), the first sheet in tab
 * order winning (F11.65). A sheet spelled exactly as the qualifier wins first, so a workbook another
 * reader loads with both `Data` and `data` keeps each one reachable. Each qualifier is folded once,
 * however often a formula looks it up, so a long qualifier in a defined name costs no more per use
 * than a short one.
 *
 * @example sheetFinder(["Data", "Main"])("Data") // "Data"
 * @example sheetFinder(["Data", "Main"])("DATA") // "Data" (case ignored)
 * @example sheetFinder(["data", "Data"])("Data") // "Data" (the exact spelling wins)
 * @example sheetFinder(["Data", "Main"])("Nope") // undefined (no such sheet)
 * @example sheetFinder([5, "Data"])("data") // "Data" (a name that is not a string is skipped)
 */
export function sheetFinder(names: readonly string[]): (qualifier: string) => string | undefined {
	const exact = new Set(names);
	const folded = new Map<string, string>();
	for (const name of names) {
		// A Workbook built without types can name a sheet with any value; no qualifier names it.
		if (typeof name !== "string") continue;
		const key = name.toLowerCase();
		if (!folded.has(key)) folded.set(key, name);
	}
	const found = new Map<string, string | undefined>();
	return (qualifier) => {
		if (found.has(qualifier)) return found.get(qualifier);
		const sheet = exact.has(qualifier) ? qualifier : folded.get(qualifier.toLowerCase());
		found.set(qualifier, sheet);
		return sheet;
	};
}

/**
 * The sheet a range lives on, from the qualifiers on its two ends (`Data!A1:A2`, `A1:data!A2`). Each
 * qualifier must name a sheet, or there is none and the range is #REF! (F11.65). The left end's sheet
 * wins, then the right end's, then the formula's own sheet.
 *
 * @example rangeSheet("Data", undefined, "Main", sheetFinder(["Data", "Main"])) // "Data"
 * @example rangeSheet(undefined, "data", "Main", sheetFinder(["Data", "Main"])) // "Data"
 * @example rangeSheet("Nope", undefined, "Main", sheetFinder(["Data", "Main"])) // undefined (#REF!)
 * @example rangeSheet("Data", "Nope", "Main", sheetFinder(["Data", "Main"])) // undefined (#REF!)
 * @example rangeSheet(undefined, undefined, "Main", sheetFinder(["Data", "Main"])) // "Main"
 */
export function rangeSheet(
	left: string | undefined,
	right: string | undefined,
	home: string,
	findSheet: (qualifier: string) => string | undefined,
): string | undefined {
	const leftSheet = left === undefined ? undefined : findSheet(left);
	if (left !== undefined && leftSheet === undefined) return undefined;
	const rightSheet = right === undefined ? undefined : findSheet(right);
	if (right !== undefined && rightSheet === undefined) return undefined;
	return leftSheet ?? rightSheet ?? home;
}

/**
 * A sheet's populated positions sorted for range queries (F11.67), both ways. `rows` lists the rows
 * that hold a cell, in ascending order, and `colsOfRow` each one's columns at the same position, in
 * ascending order. `cols` and `rowsOfCol` do the same column by column (D11).
 */
export interface SheetIndex {
	readonly rows: readonly number[];
	readonly colsOfRow: readonly (readonly number[])[];
	readonly cols: readonly number[];
	readonly rowsOfCol: readonly (readonly number[])[];
}

/**
 * The first position in an ascending array whose value is at least `value`, or the array's length
 * when none is.
 *
 * @example lowerBound([2, 4, 6], 4) // 1
 * @example lowerBound([2, 4, 6], 5) // 2
 * @example lowerBound([2, 4, 6], 7) // 3
 */
export function lowerBound(sorted: readonly number[], value: number): number {
	let lo = 0;
	let hi = sorted.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if ((sorted[mid] as number) < value) lo = mid + 1;
		else hi = mid;
	}
	return lo;
}

/**
 * The {@link SheetIndex} of a sheet's model, a map from row to a map from column to cell. A row
 * whose map is empty holds no cell, so it is left out.
 *
 * @example sheetIndex(new Map([[3, new Map([[2, "x"], [1, "y"]])], [1, new Map([[1, "z"]])]])) // { rows: [1, 3], colsOfRow: [[1], [1, 2]], cols: [1, 2], rowsOfCol: [[1, 3], [3]] }
 * @example sheetIndex(new Map([[2, new Map()]])) // { rows: [], colsOfRow: [], cols: [], rowsOfCol: [] }
 */
export function sheetIndex(cells: ReadonlyMap<number, ReadonlyMap<number, unknown>>): SheetIndex {
	const ascending = (a: number, b: number) => a - b;
	const rows = [...cells.keys()].filter((row) => (cells.get(row)?.size ?? 0) > 0).sort(ascending);
	const colsOfRow = rows.map((row) => [...(cells.get(row)?.keys() ?? [])].sort(ascending));
	// The rows are visited in ascending order, so each column's list comes out sorted.
	const byCol = new Map<number, number[]>();
	rows.forEach((row, i) => {
		for (const col of colsOfRow[i] ?? []) {
			let list = byCol.get(col);
			if (list === undefined) {
				list = [];
				byCol.set(col, list);
			}
			list.push(row);
		}
	});
	const cols = [...byCol.keys()].sort(ascending);
	const rowsOfCol = cols.map((col) => byCol.get(col) ?? []);
	return { rows, colsOfRow, cols, rowsOfCol };
}

/**
 * The populated `[col, row]` positions inside columns `c1..c2` and rows `r1..r2`, found column by
 * column and then put row by row and left to right, or `undefined` when walking the rows costs less
 * (F11.67, D11). Walking the rows visits every populated row the range spans; walking the columns
 * visits every populated column it spans, then sorts what it finds, unless it walks one column, whose
 * finds are already in order. The column walk is taken when its columns, plus its finds times the
 * sort's log factor (1 for one column), come to fewer than those rows plus the finds. A row bound
 * that is not a whole number, from a reference such as `2:3.5`, ends at the whole row below it, as
 * the row walk does.
 *
 * In the examples, `index` is `sheetIndex(new Map([[1, new Map([[1, 0]])], [3, new Map([[2, 0],
 * [5, 0]])], [9, new Map([[1, 0]])]]))`: cells at A1, B3, E3 and A9.
 *
 * @example columnWalk(index, 1, 3, 5, 3) // undefined (one row against three columns)
 * @example columnWalk(index, 1, 1, 2, 9) // undefined (two columns, but three finds to sort)
 * @example columnWalk(index, 5, 1, 5, 9) // [[5, 3]] (one column and one find, against three rows)
 * @example columnWalk(index, 4, 1, 4, 9) // [] (no populated column)
 */
export function columnWalk(
	index: SheetIndex,
	c1: number,
	r1: number,
	c2: number,
	r2: number,
): readonly (readonly [number, number])[] | undefined {
	const rowEnd = Math.floor(r2) + 1;
	const rowCount = lowerBound(index.rows, rowEnd) - lowerBound(index.rows, r1);
	const from = lowerBound(index.cols, c1);
	const to = lowerBound(index.cols, c2 + 1);
	if (to - from >= rowCount) return undefined;
	// Each column's span of rows inside the range: [column, its rows, start, end].
	const spans: (readonly [number, readonly number[], number, number])[] = [];
	let finds = 0;
	for (let k = from; k < to; k++) {
		const rows = index.rowsOfCol[k] ?? [];
		const start = lowerBound(rows, r1);
		const end = lowerBound(rows, rowEnd);
		finds += end - start;
		spans.push([index.cols[k] as number, rows, start, end]);
	}
	const one = to - from === 1;
	const sortFactor = one ? 1 : Math.ceil(Math.log2(finds + 1));
	if (to - from + finds * sortFactor >= rowCount + finds) return undefined;
	const found: (readonly [number, number])[] = [];
	for (const [col, rows, start, end] of spans) {
		for (let p = start; p < end; p++) found.push([col, rows[p] as number]);
	}
	return one ? found : found.sort((a, b) => a[1] - b[1] || a[0] - b[0]);
}

/**
 * The populated `[col, row]` positions of an index inside columns `c1..c2` and rows `r1..r2`, row by
 * row and left to right. It walks whichever is cheaper (F11.67, D11): the populated columns the range
 * spans ({@link columnWalk}), or the populated rows it spans, a binary search finding the first row
 * and, in each row, the first column. Either way the cost follows the range's own cells and the fewer
 * of its populated rows or columns, never the rest of the sheet.
 *
 * In the examples, `index` holds cells at A1, B3, E3 and A9, as in {@link columnWalk}'s.
 *
 * @example [...indexedPositions(index, 5, 1, 5, 9)] // [[5, 3]] (the column walk)
 * @example [...indexedPositions(index, 2, 1, 5, 3)] // [[2, 3], [5, 3]] (the row walk, stopping past row 3)
 * @example [...indexedPositions(index, 1, 1, 9, 99)] // [[1, 1], [2, 3], [5, 3], [1, 9]] (the row walk, to the last row)
 * @example [...indexedPositions(index, 2, 1, 4, 1)] // [] (no column in range on row 1)
 * @example [...indexedPositions(index, 1, 4, 9, 8)] // [] (no row in range)
 */
export function* indexedPositions(
	index: SheetIndex,
	c1: number,
	r1: number,
	c2: number,
	r2: number,
): Generator<readonly [number, number]> {
	const byColumn = columnWalk(index, c1, r1, c2, r2);
	if (byColumn !== undefined) {
		yield* byColumn;
		return;
	}
	for (let i = lowerBound(index.rows, r1); i < index.rows.length; i++) {
		const row = index.rows[i] as number;
		if (row > r2) return;
		const cols = index.colsOfRow[i] ?? [];
		for (let j = lowerBound(cols, c1); j < cols.length; j++) {
			const col = cols[j] as number;
			if (col > c2) break;
			yield [col, row];
		}
	}
}

/**
 * The value of one element of an array constant: a literal, or a signed literal (`-1`, `+"a"`).
 *
 * @example arrayElementValue({ type: "number", value: 2, raw: "2" }) // 2
 * @example arrayElementValue({ type: "error", value: "#N/A" }) // the #N/A error value
 * @example arrayElementValue({ type: "unary", op: "-", operand: { type: "number", value: 1, raw: "1" } }) // -1
 * @example arrayElementValue({ type: "unary", op: "-", operand: { type: "string", value: "a" } }) // #VALUE! (no number)
 * @example arrayElementValue({ type: "unary", op: "-", operand: { type: "error", value: "#DIV/0!" } }) // #DIV/0!
 * @example arrayElementValue({ type: "unary", op: "-", operand: { type: "cell", ref: "A1" } }) // #VALUE! (not a literal)
 */
export function arrayElementValue(element: ArrayElement): EvalValue {
	if (element.type === "error") return errorFromString(element.value);
	if (element.type !== "unary") return element.value;
	const operand = element.operand;
	if (operand.type === "number" || operand.type === "string" || operand.type === "boolean") {
		return applyUnary(element.op, operand.value);
	}
	if (operand.type === "error") return errorFromString(operand.value);
	return errorValue("#VALUE!");
}

// The view of each array constant node, built once: the evaluator caches the AST per formula text,
// and a view never changes (F11.64).
const arrayViews = new WeakMap<ArrayLiteral, ArrayView>();

/**
 * A function argument that is an array constant, as an {@link ArrayView} of its elements (F11.64), or
 * `undefined` for any other argument. Parentheses around the constant are looked through, and the
 * view of a node is built once.
 *
 * @example arrayArgument(parseFormula("{1,2;3,4}")) // a 2×2 ArrayView over 1, 2, 3, 4
 * @example arrayArgument(parseFormula("(({1,2}))")) // a 1×2 ArrayView over 1, 2
 * @example arrayArgument(node) === arrayArgument(node) // true: the second call reuses the view
 * @example arrayArgument(parseFormula("A1:A3")) // undefined
 */
export function arrayArgument(node: FormulaAst): ArrayView | undefined {
	let cur = node;
	while (cur.type === "group") cur = cur.expr;
	if (cur.type !== "array") return undefined;
	const cached = arrayViews.get(cur);
	if (cached !== undefined) return cached;
	const view = new ArrayView(cur.rows.map((row) => row.map(arrayElementValue)));
	arrayViews.set(cur, view);
	return view;
}

class Evaluator {
	private readonly models = new Map<string, SheetModel>();
	private readonly state = new Map<string, CellState>();
	private readonly astCache = new Map<string, FormulaAst | "error">();
	private readonly registry: FunctionRegistry;
	private readonly ctx: EvalContext;
	private readonly globalNames = new Map<string, ResolvedName>();
	private readonly scopedNames = new Map<number, Map<string, ResolvedName>>();
	// The qualifier lookup, built on first use from the sheets load() put in `models`, so a qualifier
	// never finds a sheet that did not load (F11.65). No formula runs before load().
	private sheetLookup: ((qualifier: string) => string | undefined) | undefined;
	private readonly findSheet = (qualifier: string): string | undefined => {
		this.sheetLookup ??= sheetFinder([...this.models.keys()]);
		return this.sheetLookup(qualifier);
	};
	// Each sheet's range index, built on first use from the model load() filled, which nothing
	// changes afterwards (F11.67).
	private readonly indexes = new Map<SheetModel, SheetIndex>();
	private readonly hasVolatileSource: boolean;
	private fuel: number;

	private constructor(
		private readonly workbook: Workbook,
		options: EvaluateOptions,
	) {
		this.registry = buildRegistry(options.functions);
		this.fuel = options.maxCellVisits ?? DEFAULT_MAX_VISITS;
		const now = options.now;
		const random = options.random;
		this.hasVolatileSource = now !== undefined || random !== undefined;
		// The workbook's date system threads into every serial ↔ date conversion the engine makes:
		// date-valued cells (cellToValue), the injected clock, and the date built-ins (via ctx).
		const date1904 = workbook.date1904;
		this.ctx = {
			date1904,
			now: () => {
				if (now === undefined) {
					throw new FormulaError(
						"volatile-unconfigured",
						"options.now is required for NOW/TODAY",
					);
				}
				return dateToSerial(now(), date1904);
			},
			random: () => {
				if (random === undefined) {
					throw new FormulaError(
						"volatile-unconfigured",
						"options.random is required for RAND",
					);
				}
				return random();
			},
		};
		this.indexDefinedNames();
	}

	// Build the sparse per-sheet model by draining rows() (the only enumeration of populated cells).
	private async load(): Promise<void> {
		for (let index = 0; index < this.workbook.sheets.length; index++) {
			const info = this.workbook.sheets[index];
			if (info === undefined) continue;
			const worksheet = this.workbook.sheet(info.name);
			const cells = new Map<number, Map<number, ModelCell>>();
			for await (const row of worksheet.rows()) {
				let colMap = cells.get(row.index);
				for (const cell of row.cells) {
					const at = parseA1(cell.ref);
					if (at === undefined) continue;
					const formula = worksheet.formula(cell.ref);
					if (!holdsValue(cell, formula)) {
						// A later empty copy of a ref still wins, as it does in Worksheet.cell().
						colMap?.delete(at.col);
						continue;
					}
					const model: ModelCell =
						formula !== undefined
							? { kind: "formula", text: formula }
							: { kind: "value", value: cellToValue(cell, this.workbook.date1904) };
					if (colMap === undefined) {
						colMap = new Map();
						cells.set(row.index, colMap);
					}
					colMap.set(at.col, model);
				}
			}
			this.models.set(info.name, { worksheet, index, cells });
		}
	}

	static async build(workbook: Workbook, options: EvaluateOptions): Promise<Evaluator> {
		const ev = new Evaluator(workbook, options);
		await ev.load();
		return ev;
	}

	private indexDefinedNames(): void {
		for (const dn of this.workbook.definedNames) {
			const resolved = this.classifyName(dn);
			const key = dn.name.toUpperCase();
			if (dn.localSheetId !== undefined) {
				let m = this.scopedNames.get(dn.localSheetId);
				if (m === undefined) {
					m = new Map();
					this.scopedNames.set(dn.localSheetId, m);
				}
				m.set(key, resolved);
			} else {
				this.globalNames.set(key, resolved);
			}
		}
	}

	// A defined name resolves to a constant, a simple cell/range reference, or is unsupported.
	private classifyName(dn: DefinedName): ResolvedName {
		let ast: FormulaAst;
		try {
			ast = parseFormula(dn.refersTo);
		} catch {
			return { kind: "unsupported" };
		}
		const homeSheet = this.workbook.sheets[dn.localSheetId ?? 0]?.name ?? "";
		switch (ast.type) {
			case "number":
			case "string":
			case "boolean":
				return { kind: "const", value: ast.value };
			case "error":
				return { kind: "const", value: errorFromString(ast.value) };
			case "unary":
				if (ast.op === "-" && ast.operand.type === "number")
					return { kind: "const", value: -ast.operand.value };
				if (ast.op === "+" && ast.operand.type === "number")
					return { kind: "const", value: ast.operand.value };
				return { kind: "unsupported" };
			case "cell":
			case "range":
				return { kind: "ref", ast, sheet: homeSheet };
			default:
				return { kind: "unsupported" };
		}
	}

	private lookupName(name: string, homeSheetIndex: number): ResolvedName | undefined {
		const key = name.toUpperCase();
		return this.scopedNames.get(homeSheetIndex)?.get(key) ?? this.globalNames.get(key);
	}

	// ── Cell resolution ────────────────────────────────────────────────────────────────────────

	private modelCell(sheet: string, col: number, row: number): ModelCell | undefined {
		const model = this.models.get(sheet);
		if (model === undefined) return undefined; // no such sheet → treat as absent (caller maps to #REF!)
		const cell = model.cells.get(row)?.get(col);
		if (cell !== undefined) return cell;
		// A formula-only cell (no cached value) won't be in the snapshot; recover it on demand.
		const formula = model.worksheet.formula(`${indexToColumn(col)}${row}`);
		return formula !== undefined ? { kind: "formula", text: formula } : undefined;
	}

	private setBlack(key: string, value: EvalValue): void {
		this.state.set(key, { value });
	}

	private astFor(text: string): FormulaAst | "error" {
		const cached = this.astCache.get(text);
		if (cached !== undefined) return cached;
		let result: FormulaAst | "error";
		try {
			result = parseFormula(text);
		} catch {
			result = "error";
		}
		this.astCache.set(text, result);
		return result;
	}

	/**
	 * Resolve a cell to its (memoized) value. Entry point for the top-level cell, for lazy-argument
	 * thunks, and for RangeView cells. A formula cell is driven by {@link runFormula}, which handles
	 * its whole direct-reference subtree iteratively.
	 */
	resolveCell(sheet: string, col: number, row: number, depth: number): EvalValue {
		const key = keyOf(sheet, col, row);
		const existing = this.state.get(key);
		if (existing === "grey") return errorValue("#CYCLE!");
		if (existing !== undefined) return existing.value;
		const model = this.modelCell(sheet, col, row);
		if (model === undefined) {
			this.setBlack(key, null);
			return null;
		}
		if (model.kind === "value") {
			this.setBlack(key, model.value);
			return model.value;
		}
		if (this.fuel-- <= 0) {
			throw new FormulaError("budget-exceeded", "evaluation cell-visit budget exhausted");
		}
		const ast = this.astFor(model.text);
		if (ast === "error") {
			const v = errorValue("#VALUE!");
			this.setBlack(key, v);
			return v;
		}
		this.state.set(key, "grey");
		const value = this.runFormula(key, sheet, ast, depth);
		this.setBlack(key, value);
		return value;
	}

	// The iterative driver: evaluate one formula cell and every formula cell reachable from it by
	// DIRECT references, using an explicit frame stack instead of native recursion. Each frame's
	// generator starts at `baseDepth` (the parent frame is SUSPENDED, its native frames gone), so a
	// direct-reference chain never deepens the native stack no matter how long it is.
	private runFormula(
		startKey: string,
		startSheet: string,
		startAst: FormulaAst,
		baseDepth: number,
	): EvalValue {
		if (baseDepth > MAX_NATIVE) {
			throw new FormulaError("depth-exceeded", "evaluation nesting too deep");
		}
		const stack: { key: string; gen: Generator<CellRequest, EvalValue, EvalValue> }[] = [];
		let curKey = startKey;
		let gen = this.evalGen(startAst, startSheet, baseDepth);
		let step = gen.next();
		for (;;) {
			if (step.done) {
				// A formula whose result is an empty cell holds 0, as Excel shows it (F11.62), and one whose
				// result is an array constant holds its first element (F11.64); resolveCell stores what this
				// returns. A reference range result stays a range.
				const value = blankToZero(anchorValue(step.value));
				this.setBlack(curKey, value);
				const parent = stack.pop();
				if (parent === undefined) return value;
				curKey = parent.key;
				gen = parent.gen;
				step = gen.next(value);
				continue;
			}
			const req = step.value;
			const depKey = keyOf(req.sheet, req.col, req.row);
			const s = this.state.get(depKey);
			if (s === "grey") {
				step = gen.next(errorValue("#CYCLE!"));
				continue;
			}
			if (s !== undefined) {
				step = gen.next(s.value);
				continue;
			}
			const model = this.modelCell(req.sheet, req.col, req.row);
			if (model === undefined) {
				this.setBlack(depKey, null);
				step = gen.next(null);
				continue;
			}
			if (model.kind === "value") {
				this.setBlack(depKey, model.value);
				step = gen.next(model.value);
				continue;
			}
			if (this.fuel-- <= 0) {
				throw new FormulaError("budget-exceeded", "evaluation cell-visit budget exhausted");
			}
			const depAst = this.astFor(model.text);
			if (depAst === "error") {
				const v = errorValue("#VALUE!");
				this.setBlack(depKey, v);
				step = gen.next(v);
				continue;
			}
			this.state.set(depKey, "grey");
			stack.push({ key: curKey, gen });
			curKey = depKey;
			gen = this.evalGen(depAst, req.sheet, baseDepth);
			step = gen.next();
		}
	}

	// Evaluate an AST node synchronously (for lazy-argument thunks) by running its generator and
	// resolving each requested cell through resolveCell (which itself drives iteratively). `depth`
	// is the native depth at which the thunk fires, so nested lazy calls accumulate toward the cap.
	private evalSync(node: FormulaAst, sheet: string, depth: number): EvalValue {
		const gen = this.evalGen(node, sheet, depth);
		let step = gen.next();
		while (!step.done) {
			const req = step.value;
			step = gen.next(this.resolveCell(req.sheet, req.col, req.row, depth));
		}
		return step.value;
	}

	// ── The generator evaluator ──────────────────────────────────────────────────────────────────

	private *evalGen(
		node: FormulaAst,
		sheet: string,
		depth: number,
	): Generator<CellRequest, EvalValue, EvalValue> {
		if (depth > MAX_NATIVE) {
			throw new FormulaError("depth-exceeded", "formula expression too deeply nested");
		}
		const d = depth + 1;
		switch (node.type) {
			case "number":
				return node.value;
			case "string":
				return node.value;
			case "boolean":
				return node.value;
			case "error":
				return errorFromString(node.value);
			case "empty":
				return null;
			case "array": {
				// A scalar-context array constant degrades to its top-left element (decision 5, scalar-only).
				const first = node.rows[0]?.[0];
				return first === undefined
					? errorValue("#VALUE!")
					: yield* this.evalGen(first, sheet, d);
			}
			case "group":
				return yield* this.evalGen(node.expr, sheet, d);
			case "unary": {
				const operand = yield* this.evalGen(node.operand, sheet, d);
				return applyUnary(node.op, operand);
			}
			case "postfix": {
				const operand = yield* this.evalGen(node.operand, sheet, d);
				return applyPostfix(node.op, operand);
			}
			case "binary": {
				const left = yield* this.evalGen(node.left, sheet, d);
				const right = yield* this.evalGen(node.right, sheet, d);
				return applyBinary(node.op, left, right);
			}
			case "cell": {
				// A 3-D reference (`Sheet1:Sheet3!A1`) is parse-only in v0.8 (decision 5): evaluation
				// yields a TYPED error rather than silently collapsing to the first sheet's value.
				if (node.sheet?.toName !== undefined) return errorValue("#REF!");
				const target = this.resolveSheet(node.sheet?.name, sheet);
				if (target === undefined) return errorValue("#REF!");
				const at = parseA1(node.ref);
				if (at === undefined) return errorValue("#REF!");
				return yield { sheet: target, col: at.col, row: at.row };
			}
			case "range":
				return this.resolveRange(node.left, node.right, sheet, d);
			case "name":
				if (node.sheet?.toName !== undefined) return errorValue("#REF!"); // 3-D name ref (decision 5)
				return yield* this.evalName(node.name, node.sheet?.name, sheet, d);
			case "union":
				// Multi-area unions are only meaningful to range-consuming functions (F8.3); as a scalar
				// they are #VALUE!.
				return errorValue("#VALUE!");
			case "structured":
			case "external":
				// Opaque references we do not resolve in v0.8 (decision 5) → typed error value.
				return errorValue("#REF!");
			case "call":
				return yield* this.evalCall(node.name, node.args, sheet, d);
		}
	}

	private *evalName(
		name: string,
		explicitSheet: string | undefined,
		sheet: string,
		depth: number,
	): Generator<CellRequest, EvalValue, EvalValue> {
		// `Data!Rate` looks in Data's names, then the workbook's, as `Rate` does on its own sheet; a
		// qualifier that names no sheet is #REF! (F11.65).
		const target = this.resolveSheet(explicitSheet, sheet);
		if (target === undefined) return errorValue("#REF!");
		const resolved = this.lookupName(name, this.models.get(target)?.index ?? 0);
		if (resolved === undefined) return errorValue("#NAME?");
		return yield* this.applyResolvedName(resolved, depth);
	}

	private *applyResolvedName(
		resolved: ResolvedName,
		depth: number,
	): Generator<CellRequest, EvalValue, EvalValue> {
		if (resolved.kind === "const") return resolved.value;
		if (resolved.kind === "unsupported") return errorValue("#NAME?");
		return yield* this.evalGen(resolved.ast, resolved.sheet, depth);
	}

	private *evalCall(
		name: string,
		args: readonly FormulaAst[],
		sheet: string,
		depth: number,
	): Generator<CellRequest, EvalValue, EvalValue> {
		const spec = lookupFunction(this.registry, name);
		if (spec === undefined) return errorValue("#NAME?");
		if (args.length < spec.minArgs || args.length > spec.maxArgs) return errorValue("#VALUE!");
		if (spec.volatile && !this.hasVolatileSource) {
			throw new FormulaError(
				"volatile-unconfigured",
				`${name} is volatile; inject options.now/random`,
			);
		}
		// An array constant argument reaches the function as a range of its elements (F11.64).
		if (spec.lazyArgs) {
			const thunks = args.map((arg) => {
				const array = arrayArgument(arg);
				return array === undefined ? () => this.evalSync(arg, sheet, depth) : () => array;
			});
			return spec.invoke(thunks, this.ctx);
		}
		const values: EvalValue[] = [];
		for (const arg of args) {
			const array = arrayArgument(arg);
			values.push(array === undefined ? yield* this.evalGen(arg, sheet, depth) : array);
		}
		return spec.invoke(values, this.ctx);
	}

	// ── References ────────────────────────────────────────────────────────────────────────────

	private resolveSheet(explicit: string | undefined, home: string): string | undefined {
		return explicit === undefined ? home : this.findSheet(explicit);
	}

	// Interpret a range endpoint into a partial address: a cell gives {col,row}, a column name gives
	// {col}, a row number gives {row}.
	private endpoint(node: FormulaAst): Endpoint | undefined {
		// A 3-D range endpoint (`Sheet1:Sheet3!A1`) is unsupported in v0.8 → undefined makes resolveRange
		// return a typed #REF! rather than silently binding to one sheet (decision 5).
		if ((node.type === "cell" || node.type === "name") && node.sheet?.toName !== undefined) {
			return undefined;
		}
		if (node.type === "cell") {
			const at = parseA1(node.ref);
			return at === undefined
				? undefined
				: { col: at.col, row: at.row, sheet: node.sheet?.name };
		}
		if (node.type === "name") {
			const col = COL_ONLY.exec(node.name);
			if (col !== null)
				return { col: columnToIndex(col[1] as string), sheet: node.sheet?.name };
			const row = ROW_ONLY.exec(node.name);
			if (row !== null)
				return { row: Number.parseInt(row[1] as string, 10), sheet: node.sheet?.name };
			return undefined;
		}
		if (node.type === "number") return { row: node.value };
		return undefined;
	}

	private resolveRange(
		leftNode: FormulaAst,
		rightNode: FormulaAst,
		home: string,
		depth: number,
	): EvalValue {
		const a = this.endpoint(leftNode);
		const b = this.endpoint(rightNode);
		if (a === undefined || b === undefined) return errorValue("#REF!");
		const target = rangeSheet(a.sheet, b.sheet, home, this.findSheet);
		if (target === undefined) return errorValue("#REF!");
		let startCol: number;
		let endCol: number;
		let startRow: number;
		let endRow: number;
		if (
			a.col !== undefined &&
			a.row !== undefined &&
			b.col !== undefined &&
			b.row !== undefined
		) {
			startCol = Math.min(a.col, b.col);
			endCol = Math.max(a.col, b.col);
			startRow = Math.min(a.row, b.row);
			endRow = Math.max(a.row, b.row);
		} else if (
			a.col !== undefined &&
			a.row === undefined &&
			b.col !== undefined &&
			b.row === undefined
		) {
			startCol = Math.min(a.col, b.col);
			endCol = Math.max(a.col, b.col);
			startRow = 1;
			endRow = MAX_ROW; // whole-column
		} else if (
			a.row !== undefined &&
			a.col === undefined &&
			b.row !== undefined &&
			b.col === undefined
		) {
			startRow = Math.min(a.row, b.row);
			endRow = Math.max(a.row, b.row);
			startCol = 1;
			endCol = MAX_COL; // whole-row
		} else {
			return errorValue("#REF!");
		}
		if (endCol > MAX_COL || endRow > MAX_ROW) return errorValue("#REF!");
		const model = this.models.get(target);
		const populated = () => this.populatedInRange(model, startCol, startRow, endCol, endRow);
		return new RangeView(target, startCol, startRow, endCol, endRow, populated, (col, row) =>
			this.resolveCell(target, col, row, depth),
		);
	}

	private *populatedInRange(
		model: SheetModel | undefined,
		c1: number,
		r1: number,
		c2: number,
		r2: number,
	): Generator<readonly [number, number]> {
		if (model === undefined) return;
		let index = this.indexes.get(model);
		if (index === undefined) {
			index = sheetIndex(model.cells);
			this.indexes.set(model, index);
		}
		yield* indexedPositions(index, c1, r1, c2, r2);
	}

	// ── Public results ────────────────────────────────────────────────────────────────────────

	// Reduce a top-level cell result to a scalar: a range formula shows its top-left value (decision 5).
	// The guard stops a self-referential range formula (`A1 = A1:A1`) from looping (see scalarize).
	private reduce(v: EvalValue): ScalarValue {
		let cur = v;
		for (let guard = 0; isRangeView(cur); guard++) {
			if (guard >= 64) return errorValue("#CYCLE!");
			cur = cur.topLeft();
		}
		return scalarize(cur);
	}

	evaluateCellRef(sheetName: string, ref: string): ScalarValue {
		const at = parseA1(ref);
		if (at === undefined) return errorValue("#REF!");
		// A caller without types can pass any value as the sheet name.
		if (typeof sheetName !== "string") return errorValue("#REF!");
		const sheet = this.findSheet(sheetName);
		if (sheet === undefined) return errorValue("#REF!");
		return this.reduce(this.resolveCell(sheet, at.col, at.row, 0));
	}

	evaluateAll(): WorkbookEvalResult {
		const cells: CellResult[] = [];
		const byKey = new Map<string, ScalarValue>();
		for (const [sheetName, model] of this.models) {
			for (const [row, colMap] of model.cells) {
				for (const [col, cell] of colMap) {
					if (cell.kind !== "formula") continue;
					const ref = `${indexToColumn(col)}${row}`;
					const value = this.reduce(this.resolveCell(sheetName, col, row, 0));
					cells.push({ sheet: sheetName, ref, value });
					byKey.set(`${sheetName}\u0000${ref}`, value);
				}
			}
		}
		return {
			cells,
			get: (sheet, ref) => byKey.get(`${sheet}\u0000${ref}`),
		};
	}
}

/**
 * Evaluate every formula cell in a workbook. Read-only: the workbook and its stored `<v>` caches are
 * never mutated (decision 2), so an evaluated value CAN disagree with a stale producer cache. Volatile
 * functions require `options.now`/`options.random`; a fuel budget bounds runaway inputs. Cycles resolve
 * to `#CYCLE!` values without aborting unrelated cells.
 */
export async function evaluateWorkbook(
	workbook: Workbook,
	options: EvaluateOptions = {},
): Promise<WorkbookEvalResult> {
	const evaluator = await Evaluator.build(workbook, options);
	return evaluator.evaluateAll();
}

/**
 * Evaluate a single cell (and, on demand, its transitive dependencies). Returns the scalar result — a
 * formula that yields a range is reduced to its top-left value (decision 5). The sheet name matches
 * ignoring case, as a formula's sheet qualifier does; a name that matches no sheet gives `#REF!`.
 * Throws a typed {@link FormulaError} only for configuration/resource failures (unconfigured
 * volatile, exhausted budget); in-sheet problems come back as error VALUES.
 *
 * @example await evaluateCell(wb, "Sheet1", "B2") // 84, the value of Sheet1!B2's formula
 * @example await evaluateCell(wb, "sheet1", "B2") // 84 (the sheet name ignores case)
 * @example await evaluateCell(wb, "Nope", "B2") // the #REF! error value (no such sheet)
 * @example await evaluateCell(wb, undefined, "B2") // the #REF! error value (not a sheet name)
 * @example await evaluateCell(wb, "Sheet1", "B2:C3") // the #REF! error value (not one cell)
 * @example await evaluateCell(wb, "Sheet1", "C1") // throws FormulaError when C1 calls NOW() and options.now is unset
 */
export async function evaluateCell(
	workbook: Workbook,
	sheet: string,
	ref: string,
	options: EvaluateOptions = {},
): Promise<ScalarValue> {
	const evaluator = await Evaluator.build(workbook, options);
	return evaluator.evaluateCellRef(sheet, ref);
}
