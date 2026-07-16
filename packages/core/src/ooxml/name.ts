import { isXmlSafe } from "../utils";
import { MAX_FORMULA_LEN } from "./formula";

// The defined-name-style identifier rules shared by table display names (F9.1) and workbook defined
// names (F10.1). OOXML models a table name on the same grammar as a defined name, so the legality
// check and the degrade-to-legal rewrite are single-sourced here: the tolerant reader normalizes or
// drops a foreign producer's odd name, the strict writer rejects one, and the two can never drift.

/**
 * The maximum length of a defined-name-style identifier (a table display name, a workbook defined
 * name). Single-sourced: the tolerant reader clamps/drops an over-long name here; the strict writer
 * rejects one with a typed error.
 */
export const MAX_NAME_LEN = 255;

// An identifier-shaped token that looks like a cell reference (`A1`, `ABC12`) — forbidden, along with
// the reserved bare `C`/`R` (row/column shorthand). Cell refs are ASCII, so the shape check is ASCII.
const CELL_REF_NAME = /^[A-Za-z]{1,3}[0-9]+$/;

/**
 * The ways a defined-name-style identifier can be illegal, in the order the writer checks them. A
 * legal identifier is non-empty, ≤ {@link MAX_NAME_LEN}, XML-safe, has no whitespace, starts with a
 * letter (any Unicode letter — non-English Excel locales auto-name tables `Таблица1`/`テーブル1`, F9.6)
 * / underscore / backslash, and doesn't look like a cell reference. Excel repair-prompts on a name
 * that breaks these.
 */
export type NameProblem =
	| "empty"
	| "too-long"
	| "not-xml-safe"
	| "whitespace"
	| "bad-start"
	| "cell-ref";

/**
 * The FIRST rule an identifier breaks, or `undefined` when it is legal. Single-sourced so the tolerant
 * reader and the strict writer share ONE definition of "legal name": the reader normalizes a name that
 * has a problem (F9.5), the writer rejects one — they cannot drift apart.
 */
export function nameProblem(name: string): NameProblem | undefined {
	if (name.length === 0) return "empty";
	if (name.length > MAX_NAME_LEN) return "too-long";
	if (!isXmlSafe(name)) return "not-xml-safe";
	if (/\s/.test(name)) return "whitespace";
	if (!/^[\p{L}_\\]/u.test(name)) return "bad-start";
	if (CELL_REF_NAME.test(name) || /^[CcRr]$/.test(name)) return "cell-ref";
	return undefined;
}

/**
 * Rewrite an identifier into a legal one ({@link nameProblem} of the result is `undefined`) — F9.5's
 * normalize-and-keep degradation, so a foreign name survives a read→write round-trip instead of
 * aborting the save. An already-legal name is returned UNCHANGED (a clean file stays byte-identical);
 * an illegal one is repaired deterministically: non-XML-safe chars dropped, whitespace runs collapsed
 * to `_`, a non-letter start prefixed with `_`, a cell-reference shape broken with a trailing `_`,
 * clamped to length — and `"Table"` if nothing legal survives. Used for TABLE names, whose only role
 * is a label; defined names are referenced by formulas, so they DROP rather than normalize (F10.1).
 */
export function normalizeName(name: string): string {
	if (nameProblem(name) === undefined) return name; // already legal — never touch it
	let s = "";
	for (const ch of name) {
		if (/\s/.test(ch)) {
			if (!s.endsWith("_")) s += "_";
		} else if (isXmlSafe(ch)) {
			s += ch;
		}
	}
	if (s.length > MAX_NAME_LEN) s = s.slice(0, MAX_NAME_LEN);
	if (!/^[\p{L}_\\]/u.test(s)) s = `_${s}`.slice(0, MAX_NAME_LEN);
	if (CELL_REF_NAME.test(s) || /^[CcRr]$/.test(s)) {
		s = s.length < MAX_NAME_LEN ? `${s}_` : `_${s.slice(1)}`;
	}
	return nameProblem(s) === undefined ? s : "Table";
}

// ── Defined names (F10.1) ──────────────────────────────────────────────────────────────────────────

// OOXML reserves the `_xlnm.` prefix for the BIFF built-in names (print area, auto-filter range, …).
// A user-defined name may not use it; only these specific built-in suffixes are legal with the prefix.
// Matched case-insensitively (Excel writes them in this exact casing, but a foreign producer may not).
const XLNM_PREFIX = "_xlnm.";
const XLNM_BUILTINS = new Set([
	"consolidate_area",
	"auto_open",
	"auto_close",
	"extract",
	"database",
	"criteria",
	"print_area",
	"print_titles",
	"recorder",
	"data_form",
	"auto_activate",
	"auto_deactivate",
	"sheet_title",
	"_filterdatabase",
]);

/** A defined name's legality problem: the shared identifier problems, plus a bad `_xlnm.` built-in. */
export type DefinedNameProblem = NameProblem | "bad-builtin";

/**
 * The FIRST rule a workbook defined name breaks, or `undefined` when it is legal. A defined name obeys
 * the same identifier grammar as a table name ({@link nameProblem}), with ONE addition: the reserved
 * `_xlnm.` prefix is legal only in front of a spec built-in suffix (`Print_Area`, `Print_Titles`,
 * `_FilterDatabase`, …). Single-sourced so the reader's drop and the writer's reject agree.
 */
export function definedNameProblem(name: string): DefinedNameProblem | undefined {
	// Length is the cheapest reject and it BOUNDS the work below — the `_xlnm.` branch would otherwise
	// slice/lower-case/hash an unbounded name. An over-long name is illegal on either path anyway.
	if (name.length > MAX_NAME_LEN) return "too-long";
	if (
		name.length >= XLNM_PREFIX.length &&
		name.slice(0, XLNM_PREFIX.length).toLowerCase() === XLNM_PREFIX
	) {
		const suffix = name.slice(XLNM_PREFIX.length);
		return XLNM_BUILTINS.has(suffix.toLowerCase()) ? undefined : "bad-builtin";
	}
	return nameProblem(name);
}

/**
 * Whether a parsed defined name could be re-emitted by the strict writer (F10.1, decision 3). The
 * tolerant reader DROPS a name that fails this — a foreign producer's illegal name, an empty/oversized
 * `refersTo`, or a sheet-scope pointing past the sheet list — so `Workbook.definedNames` only ever
 * holds writer-legal entries (the shared-model invariant). The writer re-checks each rule to reject
 * with a specific message; the RULES (identifier grammar, formula ceiling, XML-safety) are the same.
 * A defined name is dropped, never normalized: its name is referenced by formulas, so renaming it
 * would silently break those links.
 */
export function definedNameEmittable(
	dn: { readonly name: string; readonly refersTo: string; readonly localSheetId?: number },
	sheetCount: number,
): boolean {
	if (definedNameProblem(dn.name) !== undefined) return false;
	const r = dn.refersTo;
	// `refersTo` is a stored-form formula (no leading `=`); it must be non-empty, within Excel's
	// formula ceiling, and XML-safe to survive as element text. A hostile file can inject a control
	// char via a numeric char ref (`&#1;`), which the tokenizer decodes to an XML-illegal character.
	if (r.length === 0 || r.length > MAX_FORMULA_LEN || r.startsWith("=") || !isXmlSafe(r)) {
		return false;
	}
	if (
		dn.localSheetId !== undefined &&
		(!Number.isInteger(dn.localSheetId) || dn.localSheetId < 0 || dn.localSheetId >= sheetCount)
	) {
		return false;
	}
	return true;
}

/** The reserved built-in name Excel uses to record a sheet's autoFilter range (F10.2). */
export const FILTER_DATABASE_NAME = "_xlnm._FilterDatabase";

/**
 * Whether a defined name is the reserved `_xlnm._FilterDatabase` (case-insensitively). It is internal
 * bookkeeping OWNED by a sheet's autoFilter, not a user-facing name — the reader STRIPS it from
 * `Workbook.definedNames` (surfacing it as `Worksheet.autoFilter`) and the writer REJECTS a
 * caller-supplied one, then SYNTHESIZES it from `SheetInput.autoFilter`. Single-sourced so both agree.
 */
export function isFilterDatabaseName(name: string): boolean {
	return name.toLowerCase() === FILTER_DATABASE_NAME.toLowerCase();
}
