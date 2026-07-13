import type {
	DataValidation,
	DataValidationErrorStyle,
	DataValidationOperator,
	DataValidationType,
} from "../types";
import { isXmlSafe, localName } from "../utils";
import { tokenize } from "../xml";
import { MAX_COL, MAX_ROW, parseRef } from "./a1";

// Data-validation parser (F9.2). Reads the worksheet's MAIN-part `<dataValidations>` block into the
// shared {@link DataValidation} model. TOLERANT — and every degrade brings the value into exactly the
// set the strict writer accepts, so whatever the reader RETURNS, the writer ACCEPTS (the shared-bounds
// invariant; without it a read→modify→write of a tolerated file would abort the whole save):
//   - an unknown `type`/`operator`/`errorStyle` maps to a safe default;
//   - a non-canonical / out-of-grid `sqref` token (whole-column `A:A`, absolute `$A$1`, lowercase,
//     cross-sheet, past the grid) is DROPPED — sqref is kept symbolic, but only in the canonical form
//     the writer emits (shared `isCanonicalSqrefToken`);
//   - a prompt/error string is CLAMPED to its length bound and DROPPED if it carries an XML-unsafe
//     character (a control char or lone surrogate the writer's isXmlSafe would reject);
//   - a formula operand is taken in STORED form (a leading `=` stripped, matching the writer) and
//     dropped if empty or XML-unsafe;
//   - a rule left covering no range is dropped.
//
// Worksheet-level x14 validations (Excel 2010+ cross-sheet list sources) live under `<extLst>` as
// `<x14:dataValidations>` — the same LOCAL name as the main block. They are SKIPPED (decision 4): a
// dangling x14 twin must not round-trip. The scan tracks `<extLst>` nesting and ignores everything
// inside it, so only the main block is read.

/** Prompt/error TITLE ceiling (chars). Reader clamps, writer rejects — one shared bound (decision 5). */
export const MAX_DV_TITLE_LEN = 32;
/** Prompt/error BODY ceiling (chars), and the inline-list source ceiling. Reader clamps/drops, writer rejects. */
export const MAX_DV_TEXT_LEN = 255;
/** Max ranges in one `@sqref` — the repeat-bomb cap (decision 5). sqref stays symbolic, never per-cell. */
export const MAX_SQREF_RANGES = 65_536;

/** The eight validation types — the closed set both reader (degrade) and writer (reject) check against. */
export const DATA_VALIDATION_TYPES: readonly DataValidationType[] = [
	"none",
	"whole",
	"decimal",
	"list",
	"date",
	"time",
	"textLength",
	"custom",
];
/** The eight comparison operators. */
export const DATA_VALIDATION_OPERATORS: readonly DataValidationOperator[] = [
	"between",
	"notBetween",
	"equal",
	"notEqual",
	"lessThan",
	"lessThanOrEqual",
	"greaterThan",
	"greaterThanOrEqual",
];
/** The three error-alert styles. */
export const DATA_VALIDATION_ERROR_STYLES: readonly DataValidationErrorStyle[] = [
	"stop",
	"warning",
	"information",
];

const TYPE_SET: ReadonlySet<string> = new Set<string>(DATA_VALIDATION_TYPES);
const OPERATOR_SET: ReadonlySet<string> = new Set<string>(DATA_VALIDATION_OPERATORS);
const ERROR_STYLE_SET: ReadonlySet<string> = new Set<string>(DATA_VALIDATION_ERROR_STYLES);

// Type-predicate guards (no `as`): the sets hold exactly the union members, so membership IS the type.
// Shared by the tolerant reader (map non-members → default) and the strict writer (reject non-members).
/** True (and narrows) when `v` is one of the eight {@link DataValidationType}s. */
export function isDataValidationType(v: unknown): v is DataValidationType {
	return typeof v === "string" && TYPE_SET.has(v);
}
/** True (and narrows) when `v` is one of the eight {@link DataValidationOperator}s. */
export function isDataValidationOperator(v: unknown): v is DataValidationOperator {
	return typeof v === "string" && OPERATOR_SET.has(v);
}
/** True (and narrows) when `v` is one of the three {@link DataValidationErrorStyle}s. */
export function isDataValidationErrorStyle(v: unknown): v is DataValidationErrorStyle {
	return typeof v === "string" && ERROR_STYLE_SET.has(v);
}

// A sqref range token is a canonical uppercase A1 cell ("C1") or top-left→bottom-right range
// ("A1:B2") within Excel's grid — the SAME shape the writer emits. This is the single shared bound
// for sqref: the reader DROPS a token that fails it, the writer REJECTS one, so neither side can hand
// the other a token it refuses. `parseRef` is safe here (the regex already gated the input, and three
// letters cap the column at ZZZ, so columnToIndex can't overflow).
const CANONICAL_CELL = /^[A-Z]{1,3}[1-9][0-9]*$/;
function canonicalCell(ref: string): { readonly col: number; readonly row: number } | undefined {
	if (!CANONICAL_CELL.test(ref)) return undefined;
	const { col, row } = parseRef(ref);
	return col <= MAX_COL && row <= MAX_ROW ? { col, row } : undefined;
}
/** True when `token` is a canonical A1 cell or top-left→bottom-right range within Excel's grid. */
export function isCanonicalSqrefToken(token: string): boolean {
	const colon = token.indexOf(":");
	const from = canonicalCell(colon === -1 ? token : token.slice(0, colon));
	const to = colon === -1 ? from : canonicalCell(token.slice(colon + 1));
	if (from === undefined || to === undefined) return false;
	return to.col >= from.col && to.row >= from.row;
}

// OOXML booleans are `1`/`0` (or `true`/`false`); anything else — or an absent attribute — is undefined.
function bool01(v: string | undefined): boolean | undefined {
	if (v === "1" || v === "true") return true;
	if (v === "0" || v === "false") return false;
	return undefined;
}

// Clamp a string to at most `max` CODE POINTS. Spreading iterates by code point, so this never
// splits a surrogate pair (which would leave a lone surrogate the writer's isXmlSafe would reject —
// breaking the very bridge the shared bounds exist to keep working). Only ever fires on hostile
// input: Excel and openpyxl both enforce these limits in their own UIs.
function clampText(s: string, max: number): string {
	const cps = [...s];
	return cps.length > max ? cps.slice(0, max).join("") : s;
}

// A prompt/error string as the writer will accept it: XML-safe and within its length bound. An
// XML-unsafe value (a decoded control-character entity like `&#x1;`, or a lone surrogate) is DROPPED
// rather than clamped — there is no safe truncation of it, and the writer would reject it.
function clampSafe(raw: string, max: number): string | undefined {
	return isXmlSafe(raw) ? clampText(raw, max) : undefined;
}

// A formula operand in STORED form: a single leading `=` stripped (decision 6 — the writer stores DV
// formulas without it, and stripping here too makes read a fixpoint of write). Dropped when empty or
// XML-unsafe, so the reader never yields a formula the writer refuses.
function storedFormula(text: string): string | undefined {
	const stored = text.startsWith("=") ? text.slice(1) : text;
	return stored !== "" && isXmlSafe(stored) ? stored : undefined;
}

// A quoted inline list literal for a `list` validation (`"a,b,c"`). Its ceiling is MAX_DV_TEXT_LEN; a
// longer one is dropped on read (a mid-string clamp would break the quoting). Range/reference sources
// are unaffected.
function inlineListTooLong(type: DataValidationType, formula1: string | undefined): boolean {
	return (
		type === "list" &&
		formula1 !== undefined &&
		formula1.startsWith('"') &&
		[...formula1].length > MAX_DV_TEXT_LEN
	);
}

// The mutable rule under construction — formula1/formula2 arrive as CHILD element text.
interface Builder {
	sqref: string[];
	type: DataValidationType;
	operator?: DataValidationOperator;
	formula1?: string;
	formula2?: string;
	allowBlank?: boolean;
	showDropDown?: boolean;
	showInputMessage?: boolean;
	showErrorMessage?: boolean;
	errorStyle?: DataValidationErrorStyle;
	promptTitle?: string;
	prompt?: string;
	errorTitle?: string;
	error?: string;
}

// Split `@sqref` into individual range tokens (space-separated). Tokens are kept SYMBOLIC (never
// expanded to cells), but only in the canonical form the writer accepts — a non-canonical / out-of-grid
// token is DROPPED (shared `isCanonicalSqrefToken`). Work is capped at MAX_SQREF_RANGES tokens
// PROCESSED so a crafted whole-grid × N-thousand sqref (valid or not) can't drive unbounded validation.
function parseSqref(raw: string | undefined): string[] {
	if (raw === undefined) return [];
	const out: string[] = [];
	let processed = 0;
	for (const token of raw.split(/\s+/)) {
		if (token === "") continue;
		if (++processed > MAX_SQREF_RANGES) break;
		if (isCanonicalSqrefToken(token)) out.push(token);
	}
	return out;
}

function startBuilder(attrs: Readonly<Record<string, string>>): Builder {
	// An absent or unrecognized type degrades to "none" (tolerant).
	const type: DataValidationType = isDataValidationType(attrs.type) ? attrs.type : "none";
	const b: Builder = { sqref: parseSqref(attrs.sqref), type };

	if (isDataValidationOperator(attrs.operator)) b.operator = attrs.operator;

	const allowBlank = bool01(attrs.allowBlank);
	if (allowBlank !== undefined) b.allowBlank = allowBlank;
	// showDropDown is INVERTED in the file: `1` HIDES the arrow. Expose the intuitive boolean.
	const rawDrop = bool01(attrs.showDropDown);
	if (rawDrop !== undefined) b.showDropDown = !rawDrop;
	const showInput = bool01(attrs.showInputMessage);
	if (showInput !== undefined) b.showInputMessage = showInput;
	const showError = bool01(attrs.showErrorMessage);
	if (showError !== undefined) b.showErrorMessage = showError;

	if (isDataValidationErrorStyle(attrs.errorStyle)) b.errorStyle = attrs.errorStyle;

	// Assign a clamped+XML-safe prompt/error string only when it survives (an XML-unsafe value drops).
	const setText = (
		key: "promptTitle" | "prompt" | "errorTitle" | "error",
		raw: string | undefined,
		max: number,
	): void => {
		if (raw === undefined) return;
		const value = clampSafe(raw, max);
		if (value !== undefined) b[key] = value;
	};
	setText("promptTitle", attrs.promptTitle, MAX_DV_TITLE_LEN);
	setText("prompt", attrs.prompt, MAX_DV_TEXT_LEN);
	setText("errorTitle", attrs.errorTitle, MAX_DV_TITLE_LEN);
	setText("error", attrs.error, MAX_DV_TEXT_LEN);
	return b;
}

function finish(out: DataValidation[], b: Builder): void {
	// A rule covering no cell is meaningless (Excel never writes one) — drop it, like an empty ref.
	if (b.sqref.length === 0) return;
	// An over-long inline list literal is out of bounds and can't be safely clamped — drop the rule.
	if (inlineListTooLong(b.type, b.formula1)) return;
	const dv: DataValidation = {
		sqref: b.sqref,
		type: b.type,
		...(b.operator !== undefined ? { operator: b.operator } : {}),
		...(b.formula1 !== undefined ? { formula1: b.formula1 } : {}),
		...(b.formula2 !== undefined ? { formula2: b.formula2 } : {}),
		...(b.allowBlank !== undefined ? { allowBlank: b.allowBlank } : {}),
		...(b.showDropDown !== undefined ? { showDropDown: b.showDropDown } : {}),
		...(b.showInputMessage !== undefined ? { showInputMessage: b.showInputMessage } : {}),
		...(b.showErrorMessage !== undefined ? { showErrorMessage: b.showErrorMessage } : {}),
		...(b.errorStyle !== undefined ? { errorStyle: b.errorStyle } : {}),
		...(b.promptTitle !== undefined ? { promptTitle: b.promptTitle } : {}),
		...(b.prompt !== undefined ? { prompt: b.prompt } : {}),
		...(b.errorTitle !== undefined ? { errorTitle: b.errorTitle } : {}),
		...(b.error !== undefined ? { error: b.error } : {}),
	};
	out.push(dv);
}

/**
 * Parse a worksheet part's main `<dataValidations>` block into {@link DataValidation} rules, in
 * document order. x14 validations under `<extLst>` are skipped (decision 4). Never throws.
 */
export function parseDataValidations(xml: string): DataValidation[] {
	const out: DataValidation[] = [];
	let extLstDepth = 0; // >0 ⇒ inside an <extLst>; ignore its (x14) validations entirely
	let inBlock = false; // inside the main <dataValidations>
	let cur: Builder | undefined; // the <dataValidation> currently open
	let target: "f1" | "f2" | undefined; // which formula child is accumulating text
	let text = "";

	for (const tok of tokenize(xml)) {
		if (tok.kind === "text") {
			if (target !== undefined) text += tok.value;
			continue;
		}
		const name = localName(tok.name);
		if (tok.kind === "open") {
			if (name === "extLst") {
				if (!tok.selfClosing) extLstDepth++;
				continue;
			}
			if (extLstDepth > 0) continue;
			if (name === "dataValidations") {
				if (!tok.selfClosing) inBlock = true;
				continue;
			}
			if (!inBlock) continue;
			if (name === "dataValidation") {
				cur = startBuilder(tok.attrs);
				target = undefined;
				text = "";
				if (tok.selfClosing) {
					finish(out, cur);
					cur = undefined;
				}
				continue;
			}
			if (cur === undefined) continue;
			if (name === "formula1") {
				target = tok.selfClosing ? undefined : "f1";
				text = "";
			} else if (name === "formula2") {
				target = tok.selfClosing ? undefined : "f2";
				text = "";
			}
		} else {
			if (name === "extLst") {
				if (extLstDepth > 0) extLstDepth--;
				continue;
			}
			if (extLstDepth > 0) continue;
			if (name === "dataValidations") {
				inBlock = false;
				continue;
			}
			if (cur === undefined) continue;
			if (name === "formula1") {
				// Stored form: leading `=` stripped, empty/XML-unsafe dropped (matches the writer).
				if (target === "f1") {
					const f = storedFormula(text);
					if (f !== undefined) cur.formula1 = f;
				}
				target = undefined;
			} else if (name === "formula2") {
				if (target === "f2") {
					const f = storedFormula(text);
					if (f !== undefined) cur.formula2 = f;
				}
				target = undefined;
			} else if (name === "dataValidation") {
				finish(out, cur);
				cur = undefined;
			}
		}
	}
	return out;
}
