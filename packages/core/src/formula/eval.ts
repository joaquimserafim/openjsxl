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
	/** The evaluated value at `sheet!ref`, or `undefined` when that cell holds no formula. */
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
		if (explicitSheet !== undefined) {
			// A sheet-qualified bare name (Sheet1!Name) — resolve as a global-ish name; rare.
			const resolved = this.globalNames.get(name.toUpperCase());
			if (resolved === undefined) return errorValue("#NAME?");
			return yield* this.applyResolvedName(resolved, depth);
		}
		const model = this.models.get(sheet);
		const resolved = this.lookupName(name, model?.index ?? 0);
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
		if (explicit === undefined) return home;
		return this.models.has(explicit) ? explicit : undefined;
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
		const target = this.resolveSheet(a.sheet ?? b.sheet, home);
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
		for (const [row, colMap] of model.cells) {
			if (row < r1 || row > r2) continue;
			for (const col of colMap.keys()) {
				if (col >= c1 && col <= c2) yield [col, row];
			}
		}
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
		if (!this.models.has(sheetName)) return errorValue("#REF!");
		return this.reduce(this.resolveCell(sheetName, at.col, at.row, 0));
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
 * formula that yields a range is reduced to its top-left value (decision 5). Throws a typed
 * {@link FormulaError} only for configuration/resource failures (unconfigured volatile, exhausted
 * budget); in-sheet problems come back as error VALUES.
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
