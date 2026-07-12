import { dateToSerial, serialToDate } from "../ooxml/dates";
import { compareValues, numericStringValue, scalarize, toNumber, toText } from "./coerce";
import type {
	ArgThunk,
	EagerRegistered,
	EvalContext,
	LazyRegistered,
	RegisteredFunction,
} from "./functions";
import {
	type EvalValue,
	errorValue,
	type FormulaErrorValue,
	isErrorValue,
	isRangeView,
	type RangeView,
	type ScalarValue,
} from "./value";

// The built-in function library (F8.3). Each entry is a typed {@link RegisteredFunction}; the built-ins
// are trusted (they return an {@link EvalValue} directly, so they skip the caller-spec sanitize wrapper).
// The engine and the extension registry are proven by F8.2's caller specs; these are the batteries.
//
// Semantics follow DOCUMENTED Excel behavior, which is the authority. The Python `formulas` oracle is a
// secondary check and is KNOWN to diverge from Excel on a few points — range coercion in SUM/COUNT/
// AVERAGE (it coerces numeric text inside a reference; Excel ignores it) and TRIM (it keeps internal
// space runs; Excel collapses them). Where they differ, we match Excel and note it.
//
// One narrow, documented degradation flows from the F8.2 arg contract (locked by its tests): a
// SINGLE-CELL reference collapses to a scalar before a function sees it, so an aggregate treats it as a
// literal. `SUM(A1)` with A1 holding text/boolean therefore coerces it (Excel ignores non-numbers in a
// reference). Multi-cell ranges arrive as {@link RangeView}s and follow Excel's range rules exactly.

const POS_INF = Number.POSITIVE_INFINITY;

// ── helper constructors ─────────────────────────────────────────────────────────────────────────

function eager(
	minArgs: number,
	maxArgs: number,
	invoke: (args: readonly EvalValue[], ctx: EvalContext) => EvalValue,
	volatile = false,
): EagerRegistered {
	return { minArgs, maxArgs, volatile, lazyArgs: false, invoke };
}

function lazy(
	minArgs: number,
	maxArgs: number,
	invoke: (args: readonly ArgThunk[], ctx: EvalContext) => EvalValue,
	volatile = false,
): LazyRegistered {
	return { minArgs, maxArgs, volatile, lazyArgs: true, invoke };
}

// ── coercion helpers ────────────────────────────────────────────────────────────────────────────

// A logical test coerces numbers (0 → false, else true) and blanks (→ false); TEXT is NOT coerced
// (Excel returns #VALUE! for a text condition, verified against the oracle).
function toBool(v: EvalValue): boolean | FormulaErrorValue {
	const s = scalarize(v);
	if (isErrorValue(s)) return s;
	if (typeof s === "boolean") return s;
	if (s === null) return false;
	if (typeof s === "number") return s !== 0;
	return errorValue("#VALUE!");
}

// A number coerced to an integer index (truncated toward zero); errors propagate.
function toIndex(v: EvalValue): number | FormulaErrorValue {
	const n = toNumber(v);
	if (isErrorValue(n)) return n;
	return Math.trunc(n);
}

// A scalar rendered as a number for element-wise math, treating non-numbers (text/blank/bool) as 0;
// errors propagate. (SUMPRODUCT and friends treat text/blank as 0.)
function numericOrZero(v: EvalValue): number | FormulaErrorValue {
	const s = scalarize(v);
	if (isErrorValue(s)) return s;
	return typeof s === "number" ? s : 0;
}

function blankToZero(v: EvalValue): EvalValue {
	return v === null ? 0 : v;
}

// Excel aggregate rule: from a RANGE, only numbers participate (text/boolean/blank ignored); a SCALAR
// argument is coerced as a literal, except a blank scalar (incl. a single-cell ref to a blank cell) is
// ignored rather than counted as 0. Errors propagate.
function collectNumbers(args: readonly EvalValue[]): number[] | FormulaErrorValue {
	const out: number[] = [];
	for (const arg of args) {
		if (isRangeView(arg)) {
			for (const v of arg.values()) {
				if (isErrorValue(v)) return v;
				if (typeof v === "number") out.push(v);
			}
		} else {
			if (arg === null) continue;
			const n = toNumber(arg);
			if (isErrorValue(n)) return n;
			out.push(n);
		}
	}
	return out;
}

// Excel logical aggregate rule (AND/OR/XOR): from a range, numbers and booleans participate (text/blank
// ignored); a scalar is coerced via toBool (a text scalar → #VALUE!). Errors propagate.
function collectBools(args: readonly EvalValue[]): boolean[] | FormulaErrorValue {
	const out: boolean[] = [];
	for (const arg of args) {
		if (isRangeView(arg)) {
			for (const v of arg.values()) {
				if (isErrorValue(v)) return v;
				if (typeof v === "boolean") out.push(v);
				else if (typeof v === "number") out.push(v !== 0);
			}
		} else {
			if (arg === null) continue;
			const b = toBool(arg);
			if (isErrorValue(b)) return b;
			out.push(b);
		}
	}
	return out;
}

// Populated cells of a range in row-major order (deterministic; the model's iteration order is
// row-then-col but not guaranteed sorted, so CONCAT/TEXTJOIN sort explicitly).
function sortedEntries(rv: RangeView): { col: number; row: number; value: EvalValue }[] {
	const list = [...rv.entries()];
	list.sort((a, b) => a.row - b.row || a.col - b.col);
	return list;
}

// ── criteria (SUMIF / COUNTIF / *IFS) ─────────────────────────────────────────────────────────────

type CritOp = "=" | "<>" | ">" | "<" | ">=" | "<=";
type Criteria =
	| { readonly kind: "number"; readonly op: CritOp; readonly operand: number }
	| { readonly kind: "boolean"; readonly op: CritOp; readonly operand: boolean }
	| {
			readonly kind: "text";
			readonly op: CritOp;
			readonly operand: string;
			readonly glob: readonly GlobToken[] | null;
	  }
	| { readonly kind: "blank"; readonly op: CritOp };

function splitOp(s: string): { op: CritOp; rest: string } {
	if (s.startsWith("<=")) return { op: "<=", rest: s.slice(2) };
	if (s.startsWith(">=")) return { op: ">=", rest: s.slice(2) };
	if (s.startsWith("<>")) return { op: "<>", rest: s.slice(2) };
	if (s.startsWith("<")) return { op: "<", rest: s.slice(1) };
	if (s.startsWith(">")) return { op: ">", rest: s.slice(1) };
	if (s.startsWith("=")) return { op: "=", rest: s.slice(1) };
	return { op: "=", rest: s };
}

// Excel wildcards, matched WITHOUT a regex. A translated `*a*a*a*…` regex backtracks exponentially
// (ReDoS — a tiny hostile formula could hang the evaluator, bypassing the fuel budget), so we compile
// the pattern to tokens and match with a linear two-pointer glob algorithm: O(n·m), no backtracking.
// `*` matches any run, `?` any single char, `~` escapes the next wildcard. Anchored, case-insensitive.
type GlobToken =
	| { readonly kind: "any" }
	| { readonly kind: "one" }
	| { readonly kind: "lit"; readonly ch: string };

function parseGlob(pattern: string): GlobToken[] {
	const tokens: GlobToken[] = [];
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern.charAt(i);
		if (ch === "~") {
			const next = pattern.charAt(i + 1);
			if (next === "*" || next === "?" || next === "~") {
				tokens.push({ kind: "lit", ch: next.toUpperCase() });
				i++;
				continue;
			}
			tokens.push({ kind: "lit", ch: "~" });
			continue;
		}
		if (ch === "*") tokens.push({ kind: "any" });
		else if (ch === "?") tokens.push({ kind: "one" });
		else tokens.push({ kind: "lit", ch: ch.toUpperCase() });
	}
	return tokens;
}

function globMatch(text: string, tokens: readonly GlobToken[]): boolean {
	const up = text.toUpperCase();
	let i = 0; // text index
	let j = 0; // token index
	let star = -1; // index of the most recent '*' token
	let mark = 0; // text position captured when that '*' was taken
	while (i < up.length) {
		const t = tokens[j];
		if (t !== undefined && (t.kind === "one" || (t.kind === "lit" && t.ch === up.charAt(i)))) {
			i++;
			j++;
		} else if (t !== undefined && t.kind === "any") {
			star = j;
			mark = i;
			j++;
		} else if (star !== -1) {
			j = star + 1;
			mark++;
			i = mark;
		} else {
			return false;
		}
	}
	while (j < tokens.length && tokens[j]?.kind === "any") j++;
	return j === tokens.length;
}

function globTest(text: string, pattern: string): boolean {
	return globMatch(text, parseGlob(pattern));
}

function parseCriteria(raw: EvalValue): Criteria | FormulaErrorValue {
	const s = scalarize(raw);
	if (isErrorValue(s)) return s;
	if (typeof s === "number") return { kind: "number", op: "=", operand: s };
	if (typeof s === "boolean") return { kind: "boolean", op: "=", operand: s };
	if (s === null) return { kind: "blank", op: "=" };
	const { op, rest } = splitOp(s);
	const num = numericStringValue(rest);
	if (num !== undefined) return { kind: "number", op, operand: num };
	if (/^true$/i.test(rest)) return { kind: "boolean", op, operand: true };
	if (/^false$/i.test(rest)) return { kind: "boolean", op, operand: false };
	const glob = (op === "=" || op === "<>") && /[*?]/.test(rest) ? parseGlob(rest) : null;
	return { kind: "text", op, operand: rest, glob };
}

// Does a criterion that a blank cell can satisfy? Decides whether the *IF family must ALSO account for
// the (unpopulated) blank cells of the criteria range, which RangeView iteration never yields.
function matchesBlank(c: Criteria): boolean {
	return matchesCriteria(null, c);
}

function opSign(cmp: number, op: CritOp): boolean {
	switch (op) {
		case "=":
			return cmp === 0;
		case "<>":
			return cmp !== 0;
		case ">":
			return cmp > 0;
		case "<":
			return cmp < 0;
		case ">=":
			return cmp >= 0;
		case "<=":
			return cmp <= 0;
	}
}

function matchesCriteria(cell: EvalValue, c: Criteria): boolean {
	const v = scalarize(cell);
	if (isErrorValue(v)) return false; // an error cell never matches a normal criteria
	switch (c.kind) {
		case "number":
			if (typeof v !== "number") return false;
			return opSign(v < c.operand ? -1 : v > c.operand ? 1 : 0, c.op);
		case "boolean": {
			if (typeof v !== "boolean") return false;
			const eq = v === c.operand;
			return c.op === "<>" ? !eq : c.op === "=" ? eq : false;
		}
		case "blank": {
			const isBlank = v === null || v === "";
			return c.op === "<>" ? !isBlank : isBlank;
		}
		case "text": {
			const rendered = toText(v);
			if (isErrorValue(rendered)) return false;
			if (c.op === "=" || c.op === "<>") {
				const eq =
					c.glob !== null
						? globMatch(rendered, c.glob)
						: rendered.toUpperCase() === c.operand.toUpperCase();
				return c.op === "<>" ? !eq : eq;
			}
			const a = rendered.toUpperCase();
			const b = c.operand.toUpperCase();
			return opSign(a < b ? -1 : a > b ? 1 : 0, c.op);
		}
	}
}

// Iterate a criteria/driver argument as (value, rowOffset, colOffset) triples: a RangeView yields its
// populated cells; a scalar is a single cell at offset (0,0).
function* critCells(range: EvalValue): Generator<{ value: EvalValue; ro: number; co: number }> {
	if (isRangeView(range)) {
		for (const e of range.entries()) {
			yield { value: e.value, ro: e.row - range.startRow, co: e.col - range.startCol };
		}
	} else {
		yield { value: range, ro: 0, co: 0 };
	}
}

// Value from a companion range (sum/average range) aligned to a driver offset; a scalar companion is
// its own value only at offset (0,0). `undefined` companion means "use the driver's own value".
function alignedValue(
	range: EvalValue | undefined,
	own: EvalValue,
	ro: number,
	co: number,
): EvalValue {
	if (range === undefined) return own;
	if (isRangeView(range)) return range.cellAt(ro, co);
	return ro === 0 && co === 0 ? range : null;
}

// ── lookups ───────────────────────────────────────────────────────────────────────────────────────

// Exact-match equality for VLOOKUP/HLOOKUP/MATCH: text equality is case-insensitive and honors
// wildcards; other types compare by value (no cross-type coercion).
function lookupEquals(cell: EvalValue, lookup: ScalarValue): boolean {
	const c = scalarize(cell);
	if (isErrorValue(c)) return false;
	if (typeof lookup === "string" && /[*?]/.test(lookup) && typeof c === "string") {
		return globTest(c, lookup);
	}
	const cmp = compareValues(c, lookup);
	return !isErrorValue(cmp) && cmp === 0;
}

// The populated cells of a table's first column (VLOOKUP) or first row (HLOOKUP), sorted by offset.
function firstLine(table: RangeView, axis: "col" | "row"): { off: number; value: EvalValue }[] {
	const out: { off: number; value: EvalValue }[] = [];
	for (const e of table.entries()) {
		if (axis === "col") {
			if (e.col === table.startCol) out.push({ off: e.row - table.startRow, value: e.value });
		} else if (e.row === table.startRow) {
			out.push({ off: e.col - table.startCol, value: e.value });
		}
	}
	out.sort((a, b) => a.off - b.off);
	return out;
}

// The offset of an approximate match (largest value ≤ lookup, ascending data), or undefined.
function approxMatch(
	line: { off: number; value: EvalValue }[],
	lookup: ScalarValue,
): number | undefined {
	let best: number | undefined;
	for (const { off, value } of line) {
		const cmp = compareValues(value, lookup);
		if (isErrorValue(cmp)) continue;
		// Ascending assumption: stop at the first value that exceeds the lookup (Excel's binary-search
		// semantics — an out-of-order/too-large element ends the run rather than being skipped over).
		if (cmp <= 0) best = off;
		else break;
	}
	return best;
}

function tableLookup(args: readonly EvalValue[], axis: "col" | "row"): EvalValue {
	const lookup = scalarize(args[0] ?? null);
	if (isErrorValue(lookup)) return lookup;
	const table = args[1];
	if (!isRangeView(table)) return errorValue("#N/A");
	const idx = toIndex(args[2] ?? null);
	if (isErrorValue(idx)) return idx;
	if (idx < 1) return errorValue("#VALUE!");
	const span = axis === "col" ? table.width : table.height;
	if (idx > span) return errorValue("#REF!");
	const approx = args.length >= 4 ? toBool(args[3] ?? null) : true;
	if (isErrorValue(approx)) return approx;
	const line = firstLine(table, axis);
	if (approx) {
		const off = approxMatch(line, lookup);
		if (off === undefined) return errorValue("#N/A");
		return blankToZero(
			axis === "col" ? table.cellAt(off, idx - 1) : table.cellAt(idx - 1, off),
		);
	}
	for (const { off, value } of line) {
		if (lookupEquals(value, lookup)) {
			return blankToZero(
				axis === "col" ? table.cellAt(off, idx - 1) : table.cellAt(idx - 1, off),
			);
		}
	}
	return errorValue("#N/A");
}

// ── date helpers ────────────────────────────────────────────────────────────────────────────────

// Excel's valid date-serial range is [0, 2958465] (1899-12-30 … 9999-12-31); outside it a date
// function returns #NUM! rather than a nonsense year.
const MAX_DATE_SERIAL = 2958465;

function toSerial(v: EvalValue): number | FormulaErrorValue {
	const n = toNumber(v);
	if (isErrorValue(n)) return n;
	// A user serial may carry a fractional time-of-day; the last valid DAY is 9999-12-31 (serial
	// 2958465), so any time on that day (up to 2958465.999…) is valid — compare the truncated day.
	if (n < 0 || Math.trunc(n) > MAX_DATE_SERIAL) return errorValue("#NUM!");
	return n;
}

function daysInMonth(year: number, monthZero: number): number {
	return new Date(Date.UTC(year, monthZero + 1, 0)).getUTCDate();
}

// ── error-type map ──────────────────────────────────────────────────────────────────────────────

function errorTypeOf(code: string): number | undefined {
	switch (code) {
		case "#NULL!":
			return 1;
		case "#DIV/0!":
			return 2;
		case "#VALUE!":
			return 3;
		case "#REF!":
			return 4;
		case "#NAME?":
			return 5;
		case "#NUM!":
			return 6;
		case "#N/A":
			return 7;
		case "#GETTING_DATA":
			return 8;
		default:
			return undefined; // non-standard (e.g. #CYCLE!)
	}
}

// ── rounding ────────────────────────────────────────────────────────────────────────────────────

function roundTo(x: number, digits: number, mode: "half" | "up" | "down"): number {
	if (!Number.isFinite(x)) return x;
	const f = 10 ** Math.trunc(digits);
	if (!Number.isFinite(f)) return x; // rounding to more places than representable → unchanged
	if (f === 0) return 0; // rounding to a place coarser than the number → 0
	const y = Math.abs(x) * f;
	const r = mode === "half" ? Math.round(y) : mode === "up" ? Math.ceil(y) : Math.floor(y);
	return (Math.sign(x) || 0) * (r / f);
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════
// The registry
// ══════════════════════════════════════════════════════════════════════════════════════════════════

const ENTRIES: Record<string, RegisteredFunction> = {
	// ── math / aggregate ──────────────────────────────────────────────────────────────────────────
	SUM: eager(1, POS_INF, (args) => {
		const nums = collectNumbers(args);
		if (isErrorValue(nums)) return nums;
		let s = 0;
		for (const n of nums) s += n;
		return s;
	}),
	AVERAGE: eager(1, POS_INF, (args) => {
		const nums = collectNumbers(args);
		if (isErrorValue(nums)) return nums;
		if (nums.length === 0) return errorValue("#DIV/0!");
		let s = 0;
		for (const n of nums) s += n;
		return s / nums.length;
	}),
	COUNT: eager(1, POS_INF, (args) => {
		let n = 0;
		for (const arg of args) {
			if (isRangeView(arg)) {
				for (const v of arg.values()) if (typeof v === "number") n++;
			} else if (typeof arg === "number" || typeof arg === "boolean") {
				n++;
			} else if (typeof arg === "string" && numericStringValue(arg) !== undefined) {
				n++;
			}
			// errors, non-numeric text and blanks are ignored — COUNT never propagates an error (Excel).
		}
		return n;
	}),
	COUNTA: eager(1, POS_INF, (args) => {
		let n = 0;
		for (const arg of args) {
			if (isRangeView(arg)) n += arg.populatedCount();
			else if (arg !== null) n++;
		}
		return n;
	}),
	COUNTBLANK: eager(1, 1, (args) => {
		const arg = args[0];
		if (!isRangeView(arg)) return arg === null || arg === "" ? 1 : 0;
		let nonBlank = 0;
		for (const v of arg.values()) if (!(v === null || v === "")) nonBlank++;
		return arg.cellCount - nonBlank;
	}),
	MIN: eager(1, POS_INF, (args) => {
		const nums = collectNumbers(args);
		if (isErrorValue(nums)) return nums;
		if (nums.length === 0) return 0;
		let m = nums[0] ?? 0;
		for (const n of nums) if (n < m) m = n;
		return m;
	}),
	MAX: eager(1, POS_INF, (args) => {
		const nums = collectNumbers(args);
		if (isErrorValue(nums)) return nums;
		if (nums.length === 0) return 0;
		let m = nums[0] ?? 0;
		for (const n of nums) if (n > m) m = n;
		return m;
	}),
	MEDIAN: eager(1, POS_INF, (args) => {
		const nums = collectNumbers(args);
		if (isErrorValue(nums)) return nums;
		if (nums.length === 0) return errorValue("#NUM!");
		const s = [...nums].sort((a, b) => a - b);
		const mid = s.length >> 1;
		if (s.length % 2 === 1) return s[mid] ?? 0;
		return ((s[mid - 1] ?? 0) + (s[mid] ?? 0)) / 2;
	}),
	LARGE: eager(2, 2, (args) => rankPick(args, "large")),
	SMALL: eager(2, 2, (args) => rankPick(args, "small")),
	ROUND: eager(2, 2, (args) => numFn2(args, (x, d) => roundTo(x, d, "half"))),
	ROUNDUP: eager(2, 2, (args) => numFn2(args, (x, d) => roundTo(x, d, "up"))),
	ROUNDDOWN: eager(2, 2, (args) => numFn2(args, (x, d) => roundTo(x, d, "down"))),
	INT: eager(1, 1, (args) => numFn1(args, (x) => Math.floor(x))),
	ABS: eager(1, 1, (args) => numFn1(args, (x) => Math.abs(x))),
	SIGN: eager(1, 1, (args) => numFn1(args, (x) => Math.sign(x))),
	TRUNC: eager(1, 2, (args) => {
		const x = toNumber(args[0] ?? null);
		if (isErrorValue(x)) return x;
		const d = args.length >= 2 ? toIndex(args[1] ?? null) : 0;
		if (isErrorValue(d)) return d;
		return roundTo(x, d, "down");
	}),
	MOD: eager(2, 2, (args) => {
		const a = toNumber(args[0] ?? null);
		if (isErrorValue(a)) return a;
		const b = toNumber(args[1] ?? null);
		if (isErrorValue(b)) return b;
		if (b === 0) return errorValue("#DIV/0!");
		return a - b * Math.floor(a / b);
	}),
	POWER: eager(2, 2, (args) =>
		numFn2(args, (x, y) => (x === 0 && y < 0 ? errorValue("#DIV/0!") : x ** y)),
	),
	SQRT: eager(1, 1, (args) => numFn1(args, (x) => (x < 0 ? errorValue("#NUM!") : Math.sqrt(x)))),
	EXP: eager(1, 1, (args) => numFn1(args, (x) => Math.exp(x))),
	LN: eager(1, 1, (args) => numFn1(args, (x) => (x <= 0 ? errorValue("#NUM!") : Math.log(x)))),
	LOG10: eager(1, 1, (args) =>
		numFn1(args, (x) => (x <= 0 ? errorValue("#NUM!") : Math.log10(x))),
	),
	LOG: eager(1, 2, (args) => {
		const x = toNumber(args[0] ?? null);
		if (isErrorValue(x)) return x;
		if (x <= 0) return errorValue("#NUM!");
		const base = args.length >= 2 ? toNumber(args[1] ?? null) : 10;
		if (isErrorValue(base)) return base;
		if (base <= 0 || base === 1) return errorValue("#NUM!");
		return Math.log(x) / Math.log(base);
	}),
	PI: eager(0, 0, () => Math.PI),
	CEILING: eager(2, 2, (args) => stepRound(args, "ceil")),
	FLOOR: eager(2, 2, (args) => stepRound(args, "floor")),
	SUMPRODUCT: eager(1, POS_INF, sumproduct),

	// ── logical ───────────────────────────────────────────────────────────────────────────────────
	IF: lazy(2, 3, (thunks) => {
		const first = thunks[0];
		if (first === undefined) return errorValue("#VALUE!");
		const cond = toBool(first());
		if (isErrorValue(cond)) return cond;
		if (cond) {
			const t = thunks[1];
			return t === undefined ? true : t();
		}
		const e = thunks[2];
		return e === undefined ? false : e();
	}),
	IFERROR: lazy(2, 2, (thunks) => {
		const v = thunks[0]?.() ?? null;
		// A MULTI-cell range is not itself an error — pass it through so an aggregator around IFERROR
		// still sees the range (`SUM(IFERROR(A1:A3,0))`). `scalarize` reduces a multi-cell range to a
		// `#VALUE!` sentinel, which must NOT be mistaken for a genuine computation error. Element-wise
		// error replacement inside a range is array behaviour, out of scope for v0.8.
		if (isRangeView(v) && v.single() === undefined) return v;
		if (isErrorValue(scalarize(v))) return thunks[1]?.() ?? null;
		return v;
	}),
	IFNA: lazy(2, 2, (thunks) => {
		const v = thunks[0]?.() ?? null;
		if (isRangeView(v) && v.single() === undefined) return v; // same range pass-through as IFERROR
		const s = scalarize(v);
		if (isErrorValue(s) && s.code === "#N/A") return thunks[1]?.() ?? null;
		return v;
	}),
	IFS: lazy(2, POS_INF, (thunks) => {
		for (let i = 0; i + 1 < thunks.length; i += 2) {
			const cThunk = thunks[i];
			const vThunk = thunks[i + 1];
			if (cThunk === undefined || vThunk === undefined) break;
			const b = toBool(cThunk());
			if (isErrorValue(b)) return b;
			if (b) return vThunk();
		}
		return errorValue("#N/A");
	}),
	SWITCH: lazy(3, POS_INF, (thunks) => {
		const exprThunk = thunks[0];
		if (exprThunk === undefined) return errorValue("#VALUE!");
		const expr = scalarize(exprThunk());
		if (isErrorValue(expr)) return expr;
		let i = 1;
		for (; i + 1 < thunks.length; i += 2) {
			const caseThunk = thunks[i];
			const valThunk = thunks[i + 1];
			if (caseThunk === undefined || valThunk === undefined) break;
			const cmp = compareValues(expr, caseThunk());
			if (isErrorValue(cmp)) return cmp;
			if (cmp === 0) return valThunk();
		}
		if (i < thunks.length) {
			const d = thunks[i];
			return d === undefined ? errorValue("#N/A") : d();
		}
		return errorValue("#N/A");
	}),
	AND: eager(1, POS_INF, (args) => {
		const bools = collectBools(args);
		if (isErrorValue(bools)) return bools;
		if (bools.length === 0) return errorValue("#VALUE!");
		return bools.every((b) => b);
	}),
	OR: eager(1, POS_INF, (args) => {
		const bools = collectBools(args);
		if (isErrorValue(bools)) return bools;
		if (bools.length === 0) return errorValue("#VALUE!");
		return bools.some((b) => b);
	}),
	XOR: eager(1, POS_INF, (args) => {
		const bools = collectBools(args);
		if (isErrorValue(bools)) return bools;
		if (bools.length === 0) return errorValue("#VALUE!");
		return bools.filter((b) => b).length % 2 === 1;
	}),
	NOT: eager(1, 1, (args) => {
		const b = toBool(args[0] ?? null);
		if (isErrorValue(b)) return b;
		return !b;
	}),

	// ── lookup / reference ──────────────────────────────────────────────────────────────────────
	VLOOKUP: eager(3, 4, (args) => tableLookup(args, "col")),
	HLOOKUP: eager(3, 4, (args) => tableLookup(args, "row")),
	MATCH: eager(2, 3, matchFn),
	INDEX: eager(2, 3, indexFn),
	CHOOSE: lazy(2, POS_INF, (thunks) => {
		const idxThunk = thunks[0];
		if (idxThunk === undefined) return errorValue("#VALUE!");
		const idx = toIndex(idxThunk());
		if (isErrorValue(idx)) return idx;
		if (idx < 1 || idx >= thunks.length) return errorValue("#VALUE!");
		const chosen = thunks[idx];
		return chosen === undefined ? errorValue("#VALUE!") : chosen();
	}),
	ROWS: eager(1, 1, (args) => (isRangeView(args[0]) ? args[0].height : 1)),
	COLUMNS: eager(1, 1, (args) => (isRangeView(args[0]) ? args[0].width : 1)),

	// ── conditional aggregates ────────────────────────────────────────────────────────────────────
	SUMIF: eager(2, 3, (args) => conditionalAggregate(args[0], args[1], args[2], "sum")),
	AVERAGEIF: eager(2, 3, (args) => conditionalAggregate(args[0], args[1], args[2], "avg")),
	COUNTIF: eager(2, 2, (args) => conditionalAggregate(args[0], args[1], undefined, "count")),
	SUMIFS: eager(3, POS_INF, (args) => conditionalAggregateS(args, "sum")),
	AVERAGEIFS: eager(3, POS_INF, (args) => conditionalAggregateS(args, "avg")),
	COUNTIFS: eager(2, POS_INF, (args) => conditionalAggregateS(args, "count")),

	// ── text ──────────────────────────────────────────────────────────────────────────────────────
	CONCAT: eager(1, POS_INF, concat),
	CONCATENATE: eager(1, POS_INF, concat),
	TEXTJOIN: eager(3, POS_INF, textjoin),
	LEN: eager(1, 1, (args) => {
		const t = toText(args[0] ?? null);
		return isErrorValue(t) ? t : t.length;
	}),
	LEFT: eager(1, 2, (args) => sidePart(args, "left")),
	RIGHT: eager(1, 2, (args) => sidePart(args, "right")),
	MID: eager(3, 3, (args) => {
		const t = toText(args[0] ?? null);
		if (isErrorValue(t)) return t;
		const start = toIndex(args[1] ?? null);
		if (isErrorValue(start)) return start;
		const len = toIndex(args[2] ?? null);
		if (isErrorValue(len)) return len;
		if (start < 1 || len < 0) return errorValue("#VALUE!");
		return t.slice(start - 1, start - 1 + len);
	}),
	TRIM: eager(1, 1, (args) => {
		const t = toText(args[0] ?? null);
		if (isErrorValue(t)) return t;
		// Excel TRIM: strip leading/trailing spaces AND collapse internal runs (space char 0x20 only).
		return t
			.split(" ")
			.filter((p) => p !== "")
			.join(" ");
	}),
	UPPER: eager(1, 1, (args) => textFn1(args, (t) => t.toUpperCase())),
	LOWER: eager(1, 1, (args) => textFn1(args, (t) => t.toLowerCase())),
	PROPER: eager(1, 1, (args) => textFn1(args, proper)),
	REPT: eager(2, 2, (args) => {
		const t = toText(args[0] ?? null);
		if (isErrorValue(t)) return t;
		const k = toIndex(args[1] ?? null);
		if (isErrorValue(k)) return k;
		if (k < 0) return errorValue("#VALUE!");
		if (t.length * k > 32767) return errorValue("#VALUE!");
		return t.repeat(k);
	}),
	EXACT: eager(2, 2, (args) => {
		const a = toText(args[0] ?? null);
		if (isErrorValue(a)) return a;
		const b = toText(args[1] ?? null);
		if (isErrorValue(b)) return b;
		return a === b;
	}),
	CHAR: eager(1, 1, (args) => {
		const n = toIndex(args[0] ?? null);
		if (isErrorValue(n)) return n;
		if (n < 1 || n > 255) return errorValue("#VALUE!");
		return String.fromCharCode(n);
	}),
	CODE: eager(1, 1, (args) => {
		const t = toText(args[0] ?? null);
		if (isErrorValue(t)) return t;
		if (t.length === 0) return errorValue("#VALUE!");
		return t.charCodeAt(0);
	}),
	VALUE: eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		if (isErrorValue(s)) return s;
		if (typeof s === "number") return s;
		if (typeof s !== "string") return errorValue("#VALUE!");
		// Accept a leading currency sign, thousands commas, and a trailing percent (date-string parsing
		// is out of scope for v0.8). Everything else must be a plain en-US number.
		let t = s.trim();
		let scale = 1;
		if (t.endsWith("%")) {
			scale = 0.01;
			t = t.slice(0, -1).trim();
		}
		t = t.replace(/^\$/, "").replace(/,/g, "");
		const n = numericStringValue(t);
		return n === undefined ? errorValue("#VALUE!") : n * scale;
	}),
	SUBSTITUTE: eager(3, 4, substitute),
	REPLACE: eager(4, 4, (args) => {
		const t = toText(args[0] ?? null);
		if (isErrorValue(t)) return t;
		const start = toIndex(args[1] ?? null);
		if (isErrorValue(start)) return start;
		const num = toIndex(args[2] ?? null);
		if (isErrorValue(num)) return num;
		const nw = toText(args[3] ?? null);
		if (isErrorValue(nw)) return nw;
		if (start < 1 || num < 0) return errorValue("#VALUE!");
		return t.slice(0, start - 1) + nw + t.slice(start - 1 + num);
	}),
	FIND: eager(2, 3, (args) => findFn(args, "find")),
	SEARCH: eager(2, 3, (args) => findFn(args, "search")),

	// ── information ─────────────────────────────────────────────────────────────────────────────
	ISBLANK: eager(1, 1, (args) => scalarize(args[0] ?? null) === null),
	ISNUMBER: eager(1, 1, (args) => typeof scalarize(args[0] ?? null) === "number"),
	ISTEXT: eager(1, 1, (args) => typeof scalarize(args[0] ?? null) === "string"),
	ISLOGICAL: eager(1, 1, (args) => typeof scalarize(args[0] ?? null) === "boolean"),
	ISERROR: eager(1, 1, (args) => isErrorValue(scalarize(args[0] ?? null))),
	ISERR: eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		return isErrorValue(s) && s.code !== "#N/A";
	}),
	ISNA: eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		return isErrorValue(s) && s.code === "#N/A";
	}),
	N: eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		if (isErrorValue(s)) return s;
		if (typeof s === "number") return s;
		if (typeof s === "boolean") return s ? 1 : 0;
		return 0;
	}),
	T: eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		if (isErrorValue(s)) return s;
		return typeof s === "string" ? s : "";
	}),
	NA: eager(0, 0, () => errorValue("#N/A")),
	"ERROR.TYPE": eager(1, 1, (args) => {
		const s = scalarize(args[0] ?? null);
		if (!isErrorValue(s)) return errorValue("#N/A");
		return errorTypeOf(s.code) ?? errorValue("#N/A");
	}),

	// ── date / time ─────────────────────────────────────────────────────────────────────────────
	DATE: eager(3, 3, (args) => {
		const y = toIndex(args[0] ?? null);
		if (isErrorValue(y)) return y;
		const m = toIndex(args[1] ?? null);
		if (isErrorValue(m)) return m;
		const d = toIndex(args[2] ?? null);
		if (isErrorValue(d)) return d;
		if (y < 0 || y > 9999) return errorValue("#NUM!");
		const year = y < 1900 ? y + 1900 : y;
		const serial = dateToSerial(new Date(Date.UTC(year, m - 1, d)));
		if (!Number.isFinite(serial) || serial < 0 || serial > MAX_DATE_SERIAL) {
			return errorValue("#NUM!");
		}
		return serial;
	}),
	YEAR: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCFullYear())),
	MONTH: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCMonth() + 1)),
	DAY: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCDate())),
	HOUR: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCHours())),
	MINUTE: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCMinutes())),
	SECOND: eager(1, 1, (args) => datePart(args, (dt) => dt.getUTCSeconds())),
	WEEKDAY: eager(1, 2, (args) => {
		const serial = toSerial(args[0] ?? null);
		if (isErrorValue(serial)) return serial;
		const type = args.length >= 2 ? toIndex(args[1] ?? null) : 1;
		if (isErrorValue(type)) return type;
		const dow = serialToDate(serial).getUTCDay(); // 0=Sun..6=Sat
		if (type === 1) return dow + 1; // Sun=1..Sat=7
		if (type === 2) return ((dow + 6) % 7) + 1; // Mon=1..Sun=7
		if (type === 3) return (dow + 6) % 7; // Mon=0..Sun=6
		// Types 11..17: week starts on Mon..Sun respectively, numbered 1..7.
		if (type >= 11 && type <= 17) {
			const startDow = type === 17 ? 0 : type - 10; // 11→Mon(1) … 16→Sat(6), 17→Sun(0)
			return ((dow - startDow + 7) % 7) + 1;
		}
		return errorValue("#NUM!");
	}),
	TIME: eager(3, 3, (args) => {
		const h = toNumber(args[0] ?? null);
		if (isErrorValue(h)) return h;
		const m = toNumber(args[1] ?? null);
		if (isErrorValue(m)) return m;
		const s = toNumber(args[2] ?? null);
		if (isErrorValue(s)) return s;
		const total = h * 3600 + m * 60 + s;
		if (total < 0) return errorValue("#NUM!");
		return (total % 86400) / 86400;
	}),
	DAYS: eager(2, 2, (args) => {
		const end = toSerial(args[0] ?? null);
		if (isErrorValue(end)) return end;
		const start = toSerial(args[1] ?? null);
		if (isErrorValue(start)) return start;
		return Math.trunc(end) - Math.trunc(start);
	}),
	EDATE: eager(2, 2, (args) => monthShift(args, "same")),
	EOMONTH: eager(2, 2, (args) => monthShift(args, "end")),

	// ── volatile ──────────────────────────────────────────────────────────────────────────────────
	TODAY: eager(0, 0, (_args, ctx) => Math.floor(ctx.now()), true),
	NOW: eager(0, 0, (_args, ctx) => ctx.now(), true),
	RAND: eager(0, 0, (_args, ctx) => ctx.random(), true),
	RANDBETWEEN: eager(
		2,
		2,
		(args, ctx) => {
			const lo = toIndex(args[0] ?? null);
			if (isErrorValue(lo)) return lo;
			const hi = toIndex(args[1] ?? null);
			if (isErrorValue(hi)) return hi;
			if (lo > hi) return errorValue("#NUM!");
			return lo + Math.floor(ctx.random() * (hi - lo + 1));
		},
		true,
	),
};

// ── shared implementations referenced above ──────────────────────────────────────────────────────

function numFn1(
	args: readonly EvalValue[],
	fn: (x: number) => number | FormulaErrorValue,
): EvalValue {
	const x = toNumber(args[0] ?? null);
	if (isErrorValue(x)) return x;
	const r = fn(x);
	if (isErrorValue(r)) return r;
	if (Number.isNaN(r) || !Number.isFinite(r)) return errorValue("#NUM!"); // e.g. EXP overflow
	return r;
}

function numFn2(
	args: readonly EvalValue[],
	fn: (x: number, y: number) => number | FormulaErrorValue,
): EvalValue {
	const x = toNumber(args[0] ?? null);
	if (isErrorValue(x)) return x;
	const y = toNumber(args[1] ?? null);
	if (isErrorValue(y)) return y;
	const r = fn(x, y);
	if (isErrorValue(r)) return r;
	if (Number.isNaN(r) || !Number.isFinite(r)) return errorValue("#NUM!");
	return r;
}

function textFn1(args: readonly EvalValue[], fn: (t: string) => string): EvalValue {
	const t = toText(args[0] ?? null);
	if (isErrorValue(t)) return t;
	return fn(t);
}

function rankPick(args: readonly EvalValue[], which: "large" | "small"): EvalValue {
	const nums = collectNumbers([args[0] ?? null]);
	if (isErrorValue(nums)) return nums;
	const k = toIndex(args[1] ?? null);
	if (isErrorValue(k)) return k;
	if (k < 1 || k > nums.length) return errorValue("#NUM!");
	const sorted = [...nums].sort((a, b) => (which === "large" ? b - a : a - b));
	return sorted[k - 1] ?? errorValue("#NUM!");
}

function stepRound(args: readonly EvalValue[], mode: "ceil" | "floor"): EvalValue {
	const x = toNumber(args[0] ?? null);
	if (isErrorValue(x)) return x;
	const sig = toNumber(args[1] ?? null);
	if (isErrorValue(sig)) return sig;
	if (sig === 0) return 0;
	// Excel rejects ONLY (positive number, negative significance) with #NUM!; a negative number with a
	// positive significance is valid and rounds normally (e.g. CEILING(-2.5,1) = -2, FLOOR(-2.5,1) = -3).
	if (x > 0 && sig < 0) return errorValue("#NUM!");
	const q = x / sig;
	const r = mode === "ceil" ? Math.ceil(q) : Math.floor(q);
	return r * sig;
}

function sumproduct(args: readonly EvalValue[]): EvalValue {
	const first = args[0];
	if (!isRangeView(first)) {
		let prod = 1;
		for (const a of args) {
			const n = numericOrZero(a);
			if (isErrorValue(n)) return n;
			prod *= n;
		}
		return prod;
	}
	const h = first.height;
	const w = first.width;
	for (const a of args) {
		if (isRangeView(a)) {
			if (a.height !== h || a.width !== w) return errorValue("#VALUE!");
		} else if (h !== 1 || w !== 1) {
			return errorValue("#VALUE!");
		}
	}
	// Excel propagates an error anywhere in any array, even where the product would be 0. The sparse loop
	// below only visits positions where the first factor is non-zero, so scan every array's USED cells
	// for an error first (a blank can never be an error, so this stays bounded by populated counts).
	for (const a of args) {
		if (isRangeView(a)) {
			for (const v of a.values()) if (isErrorValue(v)) return v;
		} else if (isErrorValue(a)) {
			return a;
		}
	}
	// Sparse: only positions where the FIRST factor is non-zero can contribute (product is 0 otherwise).
	let sum = 0;
	for (const e of first.entries()) {
		const base = numericOrZero(e.value);
		if (isErrorValue(base)) return base;
		if (base === 0) continue;
		const ro = e.row - first.startRow;
		const co = e.col - first.startCol;
		let prod = base;
		let zero = false;
		for (let k = 1; k < args.length; k++) {
			const ak = args[k];
			const cell = isRangeView(ak)
				? ak.cellAt(ro, co)
				: ro === 0 && co === 0
					? (ak ?? null)
					: null;
			const n = numericOrZero(cell);
			if (isErrorValue(n)) return n;
			if (n === 0) {
				zero = true;
				break;
			}
			prod *= n;
		}
		if (!zero) sum += prod;
	}
	return sum;
}

function concat(args: readonly EvalValue[]): EvalValue {
	let out = "";
	for (const arg of args) {
		if (isRangeView(arg)) {
			for (const e of sortedEntries(arg)) {
				const t = toText(e.value);
				if (isErrorValue(t)) return t;
				out += t;
			}
		} else {
			const t = toText(arg);
			if (isErrorValue(t)) return t;
			out += t;
		}
		if (out.length > 32767) return errorValue("#VALUE!");
	}
	return out;
}

function textjoin(args: readonly EvalValue[]): EvalValue {
	const delim = toText(args[0] ?? null);
	if (isErrorValue(delim)) return delim;
	const ignoreEmpty = toBool(args[1] ?? null);
	if (isErrorValue(ignoreEmpty)) return ignoreEmpty;
	const parts: string[] = [];
	for (let i = 2; i < args.length; i++) {
		const arg = args[i];
		if (isRangeView(arg)) {
			for (const e of sortedEntries(arg)) {
				if (ignoreEmpty && (e.value === null || e.value === "")) continue;
				const t = toText(e.value);
				if (isErrorValue(t)) return t;
				parts.push(t);
			}
		} else {
			if (ignoreEmpty && (arg === null || arg === undefined || arg === "")) continue;
			const t = toText(arg ?? null);
			if (isErrorValue(t)) return t;
			parts.push(t);
		}
	}
	const joined = parts.join(delim);
	return joined.length > 32767 ? errorValue("#VALUE!") : joined;
}

function sidePart(args: readonly EvalValue[], side: "left" | "right"): EvalValue {
	const t = toText(args[0] ?? null);
	if (isErrorValue(t)) return t;
	const n = args.length >= 2 ? toIndex(args[1] ?? null) : 1;
	if (isErrorValue(n)) return n;
	if (n < 0) return errorValue("#VALUE!");
	return side === "left" ? t.slice(0, n) : n === 0 ? "" : t.slice(Math.max(0, t.length - n));
}

function proper(t: string): string {
	let out = "";
	let prevLetter = false;
	for (const ch of t) {
		const isLetter = /[A-Za-z]/.test(ch);
		out += isLetter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
		prevLetter = isLetter;
	}
	return out;
}

function substitute(args: readonly EvalValue[]): EvalValue {
	const text = toText(args[0] ?? null);
	if (isErrorValue(text)) return text;
	const oldText = toText(args[1] ?? null);
	if (isErrorValue(oldText)) return oldText;
	const newText = toText(args[2] ?? null);
	if (isErrorValue(newText)) return newText;
	if (oldText === "") return text;
	if (args.length < 4) return text.split(oldText).join(newText);
	const instance = toIndex(args[3] ?? null);
	if (isErrorValue(instance)) return instance;
	if (instance < 1) return errorValue("#VALUE!");
	let from = 0;
	let count = 0;
	for (;;) {
		const idx = text.indexOf(oldText, from);
		if (idx === -1) return text;
		count++;
		if (count === instance)
			return text.slice(0, idx) + newText + text.slice(idx + oldText.length);
		from = idx + oldText.length;
	}
}

function findFn(args: readonly EvalValue[], kind: "find" | "search"): EvalValue {
	const needle = toText(args[0] ?? null);
	if (isErrorValue(needle)) return needle;
	const hay = toText(args[1] ?? null);
	if (isErrorValue(hay)) return hay;
	const start = args.length >= 3 ? toIndex(args[2] ?? null) : 1;
	if (isErrorValue(start)) return start;
	if (start < 1) return errorValue("#VALUE!");
	const from = start - 1;
	if (from > hay.length) return errorValue("#VALUE!");
	if (kind === "find") {
		const idx = hay.indexOf(needle, from);
		return idx === -1 ? errorValue("#VALUE!") : idx + 1;
	}
	// SEARCH: case-insensitive substring. Wildcard SEARCH is deferred in v0.8 — a substring glob is an
	// O(n²) scan over attacker-controlled cell text, so `*`/`?` are matched literally here (documented).
	const idx = hay.toUpperCase().indexOf(needle.toUpperCase(), from);
	return idx === -1 ? errorValue("#VALUE!") : idx + 1;
}

function matchFn(args: readonly EvalValue[]): EvalValue {
	const lookup = scalarize(args[0] ?? null);
	if (isErrorValue(lookup)) return lookup;
	const arr = args[1];
	const type = args.length >= 3 ? toIndex(args[2] ?? null) : 1;
	if (isErrorValue(type)) return type;
	if (!isRangeView(arr)) {
		return lookupEquals(arr ?? null, lookup) ? 1 : errorValue("#N/A");
	}
	const isRow = arr.height === 1;
	const cells: { pos: number; value: EvalValue }[] = [];
	for (const e of arr.entries()) {
		cells.push({
			pos: isRow ? e.col - arr.startCol + 1 : e.row - arr.startRow + 1,
			value: e.value,
		});
	}
	cells.sort((a, b) => a.pos - b.pos);
	if (type === 0) {
		for (const { pos, value } of cells) if (lookupEquals(value, lookup)) return pos;
		return errorValue("#N/A");
	}
	// type 1: ascending data, largest value ≤ lookup. type -1: descending data, smallest value ≥ lookup.
	// Either way, stop at the first element that breaks the assumed order (Excel binary-search behavior).
	let best: number | undefined;
	for (const { pos, value } of cells) {
		const cmp = compareValues(value, lookup);
		if (isErrorValue(cmp)) continue;
		if (type >= 1) {
			if (cmp <= 0) best = pos;
			else break;
		} else {
			if (cmp >= 0) best = pos;
			else break;
		}
	}
	return best === undefined ? errorValue("#N/A") : best;
}

function indexFn(args: readonly EvalValue[]): EvalValue {
	const arr = args[0];
	const rowNum = toIndex(args[1] ?? null);
	if (isErrorValue(rowNum)) return rowNum;
	const hasCol = args.length >= 3;
	const colNum = hasCol ? toIndex(args[2] ?? null) : undefined;
	if (colNum !== undefined && isErrorValue(colNum)) return colNum;
	if (!isRangeView(arr)) {
		if (rowNum < 0 || (colNum !== undefined && colNum < 0)) return errorValue("#REF!");
		const okRow = rowNum === 1 || rowNum === 0;
		const okCol = colNum === undefined || colNum === 1 || colNum === 0;
		return okRow && okCol ? blankToZero(arr ?? null) : errorValue("#REF!");
	}
	const h = arr.height;
	const w = arr.width;
	let ro: number;
	let co: number;
	if (colNum === undefined) {
		if (h === 1) {
			ro = 0;
			co = rowNum - 1;
		} else if (w === 1) {
			ro = rowNum - 1;
			co = 0;
		} else {
			return errorValue("#REF!"); // 2D with column omitted is a whole-row array (decision 5: unsupported)
		}
	} else {
		if (rowNum === 0 || colNum === 0) return errorValue("#REF!");
		ro = rowNum - 1;
		co = colNum - 1;
	}
	if (ro < 0 || ro >= h || co < 0 || co >= w) return errorValue("#REF!");
	return blankToZero(arr.cellAt(ro, co));
}

function conditionalAggregate(
	critRange: EvalValue | undefined,
	criterion: EvalValue | undefined,
	valueRange: EvalValue | undefined,
	mode: "sum" | "avg" | "count",
): EvalValue {
	const crit = parseCriteria(criterion ?? null);
	if (isErrorValue(crit)) return crit;
	const cr = critRange ?? null;
	// When a blank cell satisfies the criterion (e.g. "<>x" or ""), the range's UNPOPULATED cells must
	// also be counted — RangeView iteration never yields them. We add them in a bounded second pass.
	const blankPass = isRangeView(cr) && matchesBlank(crit);
	const seen = blankPass && isRangeView(valueRange) ? new Set<string>() : null;
	let sum = 0;
	let count = 0;
	// 1) Populated criteria cells.
	for (const { value, ro, co } of critCells(cr)) {
		if (seen !== null) seen.add(`${ro},${co}`);
		if (!matchesCriteria(value, crit)) continue;
		if (mode === "count") {
			count++;
			continue;
		}
		const sv = alignedValue(valueRange, value, ro, co);
		if (isErrorValue(sv)) return sv;
		if (typeof sv === "number") {
			sum += sv;
			count++;
		}
	}
	// 2) Blank criteria cells (only when a blank satisfies the criterion).
	if (blankPass && isRangeView(cr)) {
		if (mode === "count") {
			count += cr.cellCount - cr.populatedCount(); // every blank position matches
		} else if (isRangeView(valueRange) && seen !== null) {
			// A blank criteria cell has no value of its own; only a POPULATED companion cell at a blank
			// position contributes. Bounded by the value range's used cells, not the rectangle.
			for (const e of valueRange.entries()) {
				const ro = e.row - valueRange.startRow;
				const co = e.col - valueRange.startCol;
				// Excel reshapes the value range to the criteria range's size (anchored top-left), so a
				// value cell OUTSIDE the criteria rectangle has no criteria position and must not count.
				if (ro >= cr.height || co >= cr.width) continue;
				if (seen.has(`${ro},${co}`)) continue; // criteria cell there is populated → already handled
				if (typeof e.value === "number") {
					sum += e.value;
					count++;
				}
			}
		}
		// SUMIF/AVERAGEIF with no separate value range: a blank criteria cell contributes no number.
	}
	if (mode === "count") return count;
	if (mode === "avg") return count === 0 ? errorValue("#DIV/0!") : sum / count;
	return sum;
}

function conditionalAggregateS(
	args: readonly EvalValue[],
	mode: "sum" | "avg" | "count",
): EvalValue {
	// SUM/AVERAGE variant: args[0] is the sum/avg range, then (critRange, criterion) pairs.
	// COUNTIFS: no leading range, just (critRange, criterion) pairs.
	const valueRange = mode === "count" ? undefined : args[0];
	const pairStart = mode === "count" ? 0 : 1;
	// A single criterion pair: delegate to the blank-aware single-criterion path so COUNTIFS/SUMIFS/
	// AVERAGEIFS agree with COUNTIF/SUMIF/AVERAGEIF on blank-satisfying criteria.
	if (mode === "count" && args.length === 2) {
		return conditionalAggregate(args[0], args[1], undefined, "count");
	}
	if (mode !== "count" && args.length === 3) {
		return conditionalAggregate(args[1], args[2], args[0], mode);
	}
	const pairs: { range: EvalValue; crit: Criteria }[] = [];
	for (let i = pairStart; i + 1 < args.length; i += 2) {
		const crit = parseCriteria(args[i + 1] ?? null);
		if (isErrorValue(crit)) return crit;
		pairs.push({ range: args[i] ?? null, crit });
	}
	if (pairs.length === 0) return errorValue("#VALUE!");
	// Choose a driver whose criterion EXCLUDES blanks: then every matching position is populated in that
	// range, so iterating its used cells (bounded) visits every match. If every criterion matches blank
	// (rare — e.g. all "<>x"), fall back to the first pair; blank positions may then be under-counted (a
	// documented v0.8 edge — the exact fix would scan the whole rectangle, which is unbounded).
	let driverIdx = pairs.findIndex((p) => !matchesBlank(p.crit));
	if (driverIdx < 0) driverIdx = 0;
	const driver = pairs[driverIdx];
	if (driver === undefined) return errorValue("#VALUE!");
	let sum = 0;
	let count = 0;
	for (const { value, ro, co } of critCells(driver.range)) {
		if (!matchesCriteria(value, driver.crit)) continue;
		let all = true;
		for (let k = 0; k < pairs.length; k++) {
			if (k === driverIdx) continue;
			const p = pairs[k];
			if (p === undefined) continue;
			if (!matchesCriteria(alignedValue(p.range, null, ro, co), p.crit)) {
				all = false;
				break;
			}
		}
		if (!all) continue;
		if (mode === "count") {
			count++;
			continue;
		}
		const sv = alignedValue(valueRange, value, ro, co);
		if (isErrorValue(sv)) return sv;
		if (typeof sv === "number") {
			sum += sv;
			count++;
		}
	}
	if (mode === "count") return count;
	if (mode === "avg") return count === 0 ? errorValue("#DIV/0!") : sum / count;
	return sum;
}

function datePart(args: readonly EvalValue[], fn: (dt: Date) => number): EvalValue {
	const serial = toSerial(args[0] ?? null);
	if (isErrorValue(serial)) return serial;
	return fn(serialToDate(serial));
}

function monthShift(args: readonly EvalValue[], mode: "same" | "end"): EvalValue {
	const serial = toSerial(args[0] ?? null);
	if (isErrorValue(serial)) return serial;
	const months = toIndex(args[1] ?? null);
	if (isErrorValue(months)) return months;
	const dt = serialToDate(serial);
	const y = dt.getUTCFullYear();
	const m = dt.getUTCMonth() + months;
	const targetYear = y + Math.floor(m / 12);
	const targetMonth = ((m % 12) + 12) % 12;
	const day =
		mode === "end"
			? daysInMonth(targetYear, targetMonth)
			: Math.min(dt.getUTCDate(), daysInMonth(targetYear, targetMonth));
	const out = dateToSerial(new Date(Date.UTC(targetYear, targetMonth, day)));
	// A result outside Excel's valid range (or a non-finite one from an absurd month count) is #NUM!.
	if (!Number.isFinite(out) || out < 0 || out > MAX_DATE_SERIAL) return errorValue("#NUM!");
	return out;
}

/** The built-in library as [name, function] entries, keyed by UPPER-CASE name. */
export const BUILTIN_ENTRIES: readonly (readonly [string, RegisteredFunction])[] =
	Object.entries(ENTRIES);
