import { BUILTIN_ENTRIES } from "./builtins";
import { FormulaError } from "./errors";
import { type EvalValue, errorValue, isErrorValue, isRangeView } from "./value";

// The function-dispatch machinery (F8.2 interface, F8.3 library). The evaluator calls into a registry
// of functions; the built-in library (SUM, IF, …) lives in `builtins.ts` and is pre-registered here. A
// caller extends the registry through `options.functions`, the same mechanism the built-ins use.

/**
 * Context passed to a function's `evaluate`. The volatile sources are the caller's injected
 * `options.now` / `options.random` (decision 3); they are only reached by functions declared
 * `volatile`, and the dispatcher rejects such a function BEFORE calling it when they are absent — so
 * inside `evaluate` they are always safe to call.
 */
export interface EvalContext {
	/** Current moment as an Excel date serial, from `options.now`. */
	now(): number;
	/** A uniform random number in [0, 1), from `options.random`. */
	random(): number;
}

/** A lazily-evaluated argument. Lazy functions (IF, CHOOSE) pull only the arguments they need. */
export type ArgThunk = () => EvalValue;

interface FunctionSpecBase {
	/** Minimum argument count (inclusive). */
	readonly minArgs: number;
	/** Maximum argument count (inclusive); use `Number.POSITIVE_INFINITY` for variadic. */
	readonly maxArgs: number;
	/** Requires `options.now`/`options.random`; rejected without them (decision 3). */
	readonly volatile?: boolean;
}

/** The common case: arguments are evaluated eagerly and handed over as values. */
export interface EagerFunctionSpec extends FunctionSpecBase {
	readonly lazyArgs?: false;
	evaluate(args: readonly EvalValue[], ctx: EvalContext): EvalValue;
}

/** IF-family functions: arguments arrive as thunks so untaken branches are never evaluated. */
export interface LazyFunctionSpec extends FunctionSpecBase {
	readonly lazyArgs: true;
	evaluate(args: readonly ArgThunk[], ctx: EvalContext): EvalValue;
}

/** A function callers register via `options.functions`, or a built-in ships as. */
export type FunctionSpec = EagerFunctionSpec | LazyFunctionSpec;

interface RegisteredBase {
	readonly minArgs: number;
	readonly maxArgs: number;
	readonly volatile: boolean;
}

/** An eager registered function: the dispatcher hands it already-evaluated argument values. */
export interface EagerRegistered extends RegisteredBase {
	readonly lazyArgs: false;
	invoke(args: readonly EvalValue[], ctx: EvalContext): EvalValue;
}

/** A lazy registered function: the dispatcher hands it thunks so untaken arguments never evaluate. */
export interface LazyRegistered extends RegisteredBase {
	readonly lazyArgs: true;
	invoke(args: readonly ArgThunk[], ctx: EvalContext): EvalValue;
}

/**
 * Normalized internal form — a discriminated union on `lazyArgs`, so the dispatcher and the built-ins
 * are fully typed with no cast: the eager branch sees `EvalValue[]`, the lazy branch sees `ArgThunk[]`.
 * A caller spec's `invoke` additionally sanitizes its result, so a stray return from an untyped caller
 * function cannot poison the evaluator (the built-ins are typed and return {@link EvalValue} directly).
 */
export type RegisteredFunction = EagerRegistered | LazyRegistered;

/** The effective function table, keyed by UPPER-CASE name (functions are case-insensitive). */
export type FunctionRegistry = ReadonlyMap<string, RegisteredFunction>;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false;
	const proto = Object.getPrototypeOf(value);
	return proto === Object.prototype || proto === null;
}

// A function may legitimately return any EvalValue (incl. a RangeView or an error value); anything
// outside that domain from an untyped caller is coerced to #VALUE! rather than trusted.
function sanitizeResult(result: unknown): EvalValue {
	if (result === null) return null;
	if (typeof result === "number" || typeof result === "string" || typeof result === "boolean") {
		return result;
	}
	if (isRangeView(result) || isErrorValue(result)) return result;
	return errorValue("#VALUE!");
}

// Validate one caller spec, reading each property exactly once (single-read TOCTOU) so a getter can't
// change the shape between validation and use. The caller's evaluate is wrapped so its result is
// sanitized and its declared arg shape (lazy vs eager) is honored by the dispatcher.
function normalizeSpec(name: string, raw: unknown): RegisteredFunction {
	if (!isPlainRecord(raw)) {
		throw new FormulaError(
			"unsupported",
			`function ${JSON.stringify(name)} is not a spec object`,
		);
	}
	const minArgs = raw.minArgs;
	const maxArgs = raw.maxArgs;
	const evaluate = raw.evaluate;
	const lazyArgs = raw.lazyArgs === true;
	const volatile = raw.volatile === true;
	if (typeof minArgs !== "number" || typeof maxArgs !== "number") {
		throw new FormulaError(
			"unsupported",
			`function ${JSON.stringify(name)} lacks numeric arg bounds`,
		);
	}
	if (typeof evaluate !== "function") {
		throw new FormulaError(
			"unsupported",
			`function ${JSON.stringify(name)} lacks an evaluate()`,
		);
	}
	// `evaluate` is a Function; calling it yields `any`, which we immediately narrow back to `unknown`
	// and sanitize — no `any` escapes, and a hostile return can't corrupt the value domain. The lazy and
	// eager branches build the matching union member so the dispatcher stays cast-free.
	if (lazyArgs) {
		return {
			minArgs,
			maxArgs,
			volatile,
			lazyArgs: true,
			invoke: (args: readonly ArgThunk[], ctx: EvalContext): EvalValue => {
				const result: unknown = evaluate(args, ctx);
				return sanitizeResult(result);
			},
		};
	}
	return {
		minArgs,
		maxArgs,
		volatile,
		lazyArgs: false,
		invoke: (args: readonly EvalValue[], ctx: EvalContext): EvalValue => {
			const result: unknown = evaluate(args, ctx);
			return sanitizeResult(result);
		},
	};
}

/**
 * Merge caller-supplied `options.functions` over the built-ins into the effective registry. Caller
 * names win and are matched case-insensitively. Each spec is validated; an invalid one is a typed
 * configuration failure, never a silent no-op.
 */
export function buildRegistry(callerFunctions: unknown): FunctionRegistry {
	const registry = new Map<string, RegisteredFunction>(BUILTIN_ENTRIES);
	if (callerFunctions === undefined) return registry;
	if (!isPlainRecord(callerFunctions)) {
		throw new FormulaError("unsupported", "options.functions must be a plain object");
	}
	for (const name of Object.keys(callerFunctions)) {
		registry.set(name.toUpperCase(), normalizeSpec(name, callerFunctions[name]));
	}
	return registry;
}
