import type { BinaryOp, PostfixOp, UnaryOp } from "./ast";
import {
	type EvalValue,
	errorValue,
	type FormulaErrorValue,
	isErrorValue,
	isRangeView,
	type ScalarValue,
} from "./value";

// The coercion + operator semantics of Excel formulas (M8 decision 6), on plain IEEE-754 doubles.
// This module is the contract the F8.2 oracle table pins cell-for-cell against Python `formulas`.
// The rules that trip people up, all encoded here:
//   • an EMPTY cell is 0 in arithmetic, "" in concat, and equals BOTH 0 and "" under `=`; an empty
//     STRING ("") is NOT a number (→ #VALUE!) — the empty-cell/empty-string distinction is real.
//   • TRUE→1 / FALSE→0 in arithmetic; numeric strings coerce (locale-invariant en-US) else #VALUE!.
//   • comparisons NEVER coerce across types and order number < text < FALSE < TRUE (text
//     case-insensitive); so `1 = "1"` is FALSE and `1 < "1"` is TRUE.
//   • errors are values and propagate left-first.

const EN_US_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * Reduce a value to a scalar: a single-cell range yields that cell; a wider range is `#VALUE!`. A
 * cell can itself evaluate to a range (`=A1:A3`), so this unwraps repeatedly until non-range.
 */
export function scalarize(v: EvalValue): ScalarValue {
	let cur = v;
	// The guard stops a degenerate self-referential single-cell range (`A1 = A1:A1`), which
	// resolves to a range pointing back at itself — cycle detection can't see it (range
	// construction is lazy), so bound the unwrap and report #VALUE! rather than loop forever.
	for (let guard = 0; isRangeView(cur); guard++) {
		if (guard >= 64) return errorValue("#VALUE!");
		const one = cur.single();
		if (one === undefined) return errorValue("#VALUE!");
		cur = one;
	}
	return cur;
}

/** Coerce a scalar to a number for arithmetic, or a typed error value when it can't be. */
export function toNumber(v: EvalValue): number | FormulaErrorValue {
	const s = scalarize(v);
	if (isErrorValue(s)) return s;
	if (s === null) return 0; // empty cell
	if (typeof s === "number") return s;
	if (typeof s === "boolean") return s ? 1 : 0;
	// A string: only a plain en-US numeric literal coerces; "" and anything else is #VALUE!.
	const t = s.trim();
	if (t !== "" && EN_US_NUMBER.test(t)) {
		const n = Number(t);
		if (Number.isFinite(n)) return n;
	}
	return errorValue("#VALUE!");
}

/** Coerce a scalar to text for concatenation. Never fails except by propagating an error value. */
export function toText(v: EvalValue): string | FormulaErrorValue {
	const s = scalarize(v);
	if (isErrorValue(s)) return s;
	if (s === null) return ""; // empty cell
	if (typeof s === "string") return s;
	if (typeof s === "boolean") return s ? "TRUE" : "FALSE";
	return numberToText(s);
}

/**
 * Excel "General" number→text for concatenation: exact for integers, 15 significant digits (Excel's
 * precision) with trailing zeros trimmed otherwise. Full General rendering (scientific thresholds,
 * width-dependent rounding) is the number-format renderer's job — out of F8.2 scope (the TEXT
 * function is F8.3-excluded for the same reason) — so exotic magnitudes may differ from Excel here.
 */
function numberToText(n: number): string {
	if (n === 0) return "0"; // also normalizes -0
	if (Number.isInteger(n) && Math.abs(n) < 1e16) return n.toString();
	const p = n.toPrecision(15);
	if (/[eE]/.test(p)) return p;
	return p.includes(".") ? p.replace(/0+$/, "").replace(/\.$/, "") : p;
}

// Type rank for cross-type comparison ordering: number < text < boolean (decision 6).
function typeRank(v: number | string | boolean): 0 | 1 | 2 {
	if (typeof v === "number") return 0;
	if (typeof v === "string") return 1;
	return 2;
}

// What an empty cell becomes when compared against a value of a given type.
function emptyAs(other: number | string | boolean): number | string | boolean {
	if (typeof other === "number") return 0;
	if (typeof other === "string") return "";
	return false;
}

/**
 * Compare two scalars by Excel's rules for the relational operators: no cross-type coercion, ordering
 * number < text < FALSE < TRUE, text case-insensitive. Returns -1/0/1, or a propagated error value.
 * An empty operand takes the other's type (so empty equals 0, "", and FALSE).
 */
export function compareValues(a: EvalValue, b: EvalValue): -1 | 0 | 1 | FormulaErrorValue {
	const x0 = scalarize(a);
	if (isErrorValue(x0)) return x0;
	const y0 = scalarize(b);
	if (isErrorValue(y0)) return y0;
	// Empty takes the other operand's type; two empties are equal.
	if (x0 === null && y0 === null) return 0;
	if (x0 === null) return y0 === null ? 0 : compareScalars(emptyAs(y0), y0);
	if (y0 === null) return compareScalars(x0, emptyAs(x0));
	return compareScalars(x0, y0);
}

function compareScalars(x: number | string | boolean, y: number | string | boolean): -1 | 0 | 1 {
	const rx = typeRank(x);
	const ry = typeRank(y);
	if (rx !== ry) return rx < ry ? -1 : 1;
	if (typeof x === "number" && typeof y === "number") return sign(x - y);
	if (typeof x === "boolean" && typeof y === "boolean") return sign(Number(x) - Number(y));
	if (typeof x === "string" && typeof y === "string") {
		// Excel compares text case-insensitively.
		const sa = x.toUpperCase();
		const sb = y.toUpperCase();
		return sa < sb ? -1 : sa > sb ? 1 : 0;
	}
	return 0; // unreachable: equal type ranks imply same primitive type
}

function sign(n: number): -1 | 0 | 1 {
	return n < 0 ? -1 : n > 0 ? 1 : 0;
}

function arithmetic(op: "+" | "-" | "*" | "/" | "^", a: number, b: number): EvalValue {
	let result: number;
	switch (op) {
		case "+":
			result = a + b;
			break;
		case "-":
			result = a - b;
			break;
		case "*":
			result = a * b;
			break;
		case "/":
			if (b === 0) return errorValue("#DIV/0!");
			result = a / b;
			break;
		case "^":
			// 0 to a negative power is a division by zero; a negative base to a fractional power is
			// non-real. Everything else that overflows to ±Infinity or NaN is #NUM!.
			if (a === 0 && b < 0) return errorValue("#DIV/0!");
			result = a ** b;
			break;
	}
	if (Number.isNaN(result)) return errorValue("#NUM!");
	if (!Number.isFinite(result)) return errorValue("#NUM!");
	return result;
}

/** Apply an infix operator to two already-evaluated operands. */
export function applyBinary(op: BinaryOp["op"], left: EvalValue, right: EvalValue): EvalValue {
	if (op === "&") {
		const x = toText(left);
		if (isErrorValue(x)) return x;
		const y = toText(right);
		if (isErrorValue(y)) return y;
		return x + y;
	}
	if (op === "=" || op === "<>" || op === "<" || op === ">" || op === "<=" || op === ">=") {
		const c = compareValues(left, right);
		if (isErrorValue(c)) return c;
		switch (op) {
			case "=":
				return c === 0;
			case "<>":
				return c !== 0;
			case "<":
				return c < 0;
			case ">":
				return c > 0;
			case "<=":
				return c <= 0;
			default:
				return c >= 0;
		}
	}
	const x = toNumber(left);
	if (isErrorValue(x)) return x;
	const y = toNumber(right);
	if (isErrorValue(y)) return y;
	return arithmetic(op, x, y);
}

/** Apply a prefix operator (`-`, `+`, or the `@` implicit-intersection reducer). */
export function applyUnary(op: UnaryOp["op"], operand: EvalValue): EvalValue {
	if (op === "@") return scalarize(operand);
	const n = toNumber(operand);
	if (isErrorValue(n)) return n;
	return op === "-" ? -n : n;
}

/** Apply a postfix operator (`%` percent, `#` spilled-range). */
export function applyPostfix(op: PostfixOp["op"], operand: EvalValue): EvalValue {
	if (op === "%") {
		const n = toNumber(operand);
		if (isErrorValue(n)) return n;
		return n / 100;
	}
	// `#` (spilled range): dynamic-array spills are not evaluated in v0.8, so the spill of an anchor
	// is unknowable — a named degradation to #REF! rather than a silently-wrong single value.
	return errorValue("#REF!");
}
