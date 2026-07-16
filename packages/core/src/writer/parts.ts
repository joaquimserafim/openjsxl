import { XlsxError } from "../errors";
import { type CellRef, indexToColumn, parseCanonicalRange } from "../ooxml/a1";
import { MEDIA_MIME_TO_EXT } from "../ooxml/drawing";
import { MAX_FORMULA_LEN } from "../ooxml/formula";
import {
	type DefinedNameProblem,
	definedNameProblem,
	FILTER_DATABASE_NAME,
	isFilterDatabaseName,
	MAX_NAME_LEN,
} from "../ooxml/name";
import type { DefinedName } from "../ooxml/workbook";
import { encodeXstring } from "../ooxml/xstring";
import type { SheetState } from "../types";
import type { SheetRel, SheetSideParts } from "./sheet";
import { DEFAULT_THEME_XML } from "./theme";
import { escapeAttr, escapeText, isPlainRecord, isXmlSafe } from "./xml";

// The OPC part builders shared by the buffered writer (writeXlsx) and the streaming writer
// (streamXlsx). Every function here returns the EXACT bytes writeXlsx used to build inline, so the
// buffered writer's golden pins are untouched; the streaming writer reuses them verbatim.

const encoder = new TextEncoder();
export const encode = (xml: string): Uint8Array => encoder.encode(xml);

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const CT_BASE = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml";

// Excel's sheet-name rules: 1–31 chars, none of \ / ? * [ ] :, and unique ignoring case. A file that
// breaks these opens with a "repair" prompt, so we refuse to write it.
const MAX_SHEET_NAME = 31;
const FORBIDDEN_SHEET_NAME = /[\\/?*[\]:]/;

/**
 * The writer rejects invalid input with a typed error rather than crashing — so before either entry
 * point reads `workbook.sheets`, confirm the workbook is a non-null object. A `null`/`undefined`
 * workbook would otherwise throw a raw `TypeError` on that property access (found by the F9.4 writer
 * fuzzer); `sheets` itself is then validated by {@link validateSheetMeta}. Single-sourced so both the
 * buffered and streaming entries reject identically.
 */
export function requireWorkbookObject(workbook: unknown): void {
	if (typeof workbook !== "object" || workbook === null) {
		throw new XlsxError("invalid-input", "a workbook must be an object with a `sheets` array");
	}
}

/**
 * Validate every sheet's NAME and visibility STATE (the checks identical to both writers) and return
 * the resolved names and states. Each `name`/`state` is read exactly once (getters/Proxies can vary
 * between reads), so downstream emission uses these arrays, never `sheet.name`/`sheet.state` — closing
 * a TOCTOU gap where a flipped getter could otherwise slip a forbidden name into workbook.xml. Does
 * NOT validate `rows` — the buffered and streaming writers accept different row shapes (an array vs an
 * iterable) and each checks its own.
 */
export function validateSheetMeta(sheets: unknown): {
	states: SheetState[];
	names: string[];
} {
	if (!Array.isArray(sheets) || sheets.length === 0) {
		throw new XlsxError("invalid-input", "a workbook needs at least one sheet");
	}
	const seen = new Set<string>();
	const states: SheetState[] = [];
	const names: string[] = [];
	let anyVisible = false;
	for (const sheet of sheets) {
		const name = sheet?.name;
		if (typeof name !== "string" || name.length === 0) {
			throw new XlsxError("invalid-input", "a sheet name must be a non-empty string");
		}
		if (name.length > MAX_SHEET_NAME) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" exceeds ${MAX_SHEET_NAME} characters`,
			);
		}
		if (FORBIDDEN_SHEET_NAME.test(name)) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" contains a forbidden character (\\ / ? * [ ] :)`,
			);
		}
		// A control character or lone surrogate in the name would corrupt the workbook.xml attribute
		// the same way it would a cell — the whole file would then fail to open.
		if (!isXmlSafe(name)) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" contains a character not allowed in XML`,
			);
		}
		// Excel treats sheet names case-insensitively, so "Data" and "data" collide.
		const key = name.toLowerCase();
		if (seen.has(key)) {
			throw new XlsxError(
				"invalid-input",
				`duplicate sheet name "${name}" (case-insensitive)`,
			);
		}
		seen.add(key);
		names.push(name);
		// Tab visibility (F4.6). Read the caller's property ONCE; absent means visible; only the three
		// spec values are accepted.
		const state = sheet.state ?? "visible";
		if (state !== "visible" && state !== "hidden" && state !== "veryHidden") {
			throw new XlsxError(
				"invalid-input",
				`sheet "${name}": state must be "visible", "hidden", or "veryHidden"`,
			);
		}
		states.push(state);
		if (state === "visible") anyVisible = true;
	}
	// Excel refuses a workbook whose every sheet is hidden — there would be nothing to show.
	if (!anyVisible) {
		throw new XlsxError("invalid-input", "at least one sheet must be visible");
	}
	return { states, names };
}

// The content type for each media extension the picture writer may emit (F6.3) — DERIVED by
// inverting the canonical writer allowlist (ooxml/drawing.ts), so the Default entries here can
// never drift from the extensions the media registry actually names its parts with.
const MEDIA_EXT_TO_MIME: Readonly<Record<string, string>> = Object.fromEntries(
	Object.entries(MEDIA_MIME_TO_EXT).map(([mime, ext]) => [ext, mime]),
);

/**
 * `[Content_Types].xml` — one Override per part not covered by the rels/xml defaults. The `vml`
 * Default is emitted iff any sheet has comments, and image `Default`s + drawing `Override`s iff any
 * sheet has pictures — all derived from the actual `commentSheets`/`drawingSheets`/`mediaExtensions`
 * so no caller can pass a flag inconsistent with the parts really written. When there are no comments
 * and no images the output is byte-identical to the pre-F6.3 map.
 */
export function contentTypesXml(
	sheetCount: number,
	needStyles: boolean,
	needTheme: boolean,
	commentSheets: readonly number[],
	drawingSheets: readonly number[] = [],
	mediaExtensions: readonly string[] = [],
	tablePartNumbers: readonly number[] = [],
): string {
	const needVml = commentSheets.length > 0;
	const overrides = [
		`<Override PartName="/xl/workbook.xml" ContentType="${CT_BASE}.sheet.main+xml"/>`,
		...Array.from(
			{ length: sheetCount },
			(_, i) =>
				`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT_BASE}.worksheet+xml"/>`,
		),
		...(needStyles
			? [`<Override PartName="/xl/styles.xml" ContentType="${CT_BASE}.styles+xml"/>`]
			: []),
		...(needTheme
			? [
					'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
				]
			: []),
		...commentSheets.map(
			(i) =>
				`<Override PartName="/xl/comments${i + 1}.xml" ContentType="${CT_BASE}.comments+xml"/>`,
		),
		...drawingSheets.map(
			(i) =>
				`<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
		),
		...tablePartNumbers.map(
			(n) =>
				`<Override PartName="/xl/tables/table${n}.xml" ContentType="${CT_BASE}.table+xml"/>`,
		),
	].join("");
	const vmlDefault = needVml
		? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
		: "";
	const mediaDefaults = mediaExtensions
		.map((ext) => `<Default Extension="${ext}" ContentType="${MEDIA_EXT_TO_MIME[ext]}"/>`)
		.join("");
	return `${XML_DECL}\n<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="${CT_RELS}"/><Default Extension="xml" ContentType="application/xml"/>${vmlDefault}${mediaDefaults}${overrides}</Types>`;
}

/** `_rels/.rels` — the package's single relationship: officeDocument → the workbook. */
export function packageRelsXml(): string {
	return `${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
}

/**
 * `xl/workbook.xml` — the sheet list, with visibility state and the active-tab fix-up. `names` are the
 * VALIDATED names from {@link validateSheetMeta} (not re-read from the caller), so a getter can't slip
 * a different name in between validation and emission. `definedNames` are the VALIDATED entries from
 * {@link validateDefinedNames}; they emit in the load-bearing CT_Workbook slot right after `<sheets>`.
 */
export function workbookXml(
	names: readonly string[],
	states: readonly SheetState[],
	date1904: boolean,
	definedNames: readonly DefinedName[] = [],
): string {
	const workbookPr = date1904 ? '<workbookPr date1904="1"/>' : "";
	// The active tab defaults to index 0; when the FIRST sheet is hidden that default would point at a
	// tab the user can't see, so aim it at the first visible sheet instead — exactly what openpyxl
	// does. All-visible workbooks emit no <bookViews> and keep their pre-F4.6 bytes.
	const firstVisible = states.indexOf("visible");
	const bookViews =
		firstVisible > 0
			? `<bookViews><workbookView activeTab="${firstVisible}"/></bookViews>`
			: "";
	const sheetsXml = names
		.map((name, i) => {
			const state = states[i] === "visible" ? "" : ` state="${states[i]}"`;
			return `<sheet name="${escapeAttr(name)}" sheetId="${i + 1}"${state} r:id="rId${i + 1}"/>`;
		})
		.join("");
	return `${XML_DECL}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${workbookPr}${bookViews}<sheets>${sheetsXml}</sheets>${definedNamesXml(definedNames)}</workbook>`;
}

// <definedNames> — one <definedName> per already-validated entry, emitted in the CT_Workbook slot
// right after <sheets> (element order is load-bearing; Excel repair-prompts on a violation). Absent
// (no element at all) when there are none, so a names-free workbook keeps its exact pre-F10.1 bytes.
// Entries are trusted here — validateDefinedNames has already checked and single-read them.
function definedNamesXml(names: readonly DefinedName[]): string {
	if (names.length === 0) return "";
	const items = names
		.map((dn) => {
			// CT_DefinedName attribute order: name, …, localSheetId, hidden. @name is ST_Xstring (the
			// same type as a table displayName, F9.6), so encode it — a name like `_x0041_` survives Excel's
			// decode. refersTo is element TEXT (a stored-form formula), escaped like cell-formula text.
			const scope = dn.localSheetId !== undefined ? ` localSheetId="${dn.localSheetId}"` : "";
			const hidden = dn.hidden ? ' hidden="1"' : "";
			return `<definedName name="${escapeAttr(encodeXstring(dn.name))}"${scope}${hidden}>${escapeText(dn.refersTo)}</definedName>`;
		})
		.join("");
	return `<definedNames>${items}</definedNames>`;
}

// Human message for an illegal defined name, by problem code (the writer names WHY it rejected).
function definedNameProblemMessage(name: string, problem: DefinedNameProblem): string {
	switch (problem) {
		case "empty":
			return "must not be empty";
		case "too-long":
			return `"${name}" exceeds ${MAX_NAME_LEN} characters`;
		case "not-xml-safe":
			return `"${name}" contains a character not allowed in XML`;
		case "whitespace":
			return `"${name}" must not contain whitespace`;
		case "bad-start":
			return `"${name}" must start with a letter, underscore, or backslash`;
		case "cell-ref":
			return `"${name}" must not look like a cell reference`;
		case "bad-builtin":
			return `"${name}" uses the reserved "_xlnm." prefix but is not a spec built-in name`;
	}
}

/**
 * Validate `WorkbookInput.definedNames` (F10.1) into the trusted array {@link workbookXml} emits — or
 * `[]` when absent, so a names-free workbook stays byte-identical. Each entry's properties are read
 * exactly ONCE into locals (single-read TOCTOU) before validating and building the returned object, so
 * a getter/Proxy can't vary a value between check and emission. Rejects — with a typed {@link
 * XlsxError} — anything the writer can't emit: a non-array, a non-plain-object entry, an unknown
 * property, an illegal name (the shared defined-name grammar + the reserved `_xlnm.` prefix), a
 * `refersTo` that isn't a non-empty stored-form (no leading `=`) XML-safe formula within Excel's
 * ceiling, a `localSheetId` that isn't an integer index of an existing sheet, a non-boolean `hidden`,
 * or a duplicate name within one scope (case-insensitive, exactly as Excel treats them). The rules
 * mirror the reader's `definedNameEmittable` drop test — same bounds, opposite response.
 */
export function validateDefinedNames(raw: unknown, sheetCount: number): DefinedName[] {
	if (raw === undefined) return [];
	if (!Array.isArray(raw)) {
		throw new XlsxError("invalid-input", "definedNames must be an array");
	}
	const out: DefinedName[] = [];
	// Duplicate detection per scope: key = `${scopeKey} ${NAME_UPPERCASE}`. scopeKey is "*" (global) or
	// the sheet index; a legal name has no whitespace (nameProblem rejects it), so the single-space
	// separator can never merge two distinct scope/name pairs.
	const seen = new Set<string>();
	for (let i = 0; i < raw.length; i++) {
		const entry = raw[i];
		if (!isPlainRecord(entry)) {
			throw new XlsxError("invalid-input", `definedNames[${i}] must be an object`);
		}
		for (const key of Object.keys(entry)) {
			if (
				key !== "name" &&
				key !== "refersTo" &&
				key !== "localSheetId" &&
				key !== "hidden"
			) {
				throw new XlsxError(
					"invalid-input",
					`definedNames[${i}] has an unknown property "${key}"`,
				);
			}
		}
		const name = entry.name;
		const refersTo = entry.refersTo;
		const localSheetId = entry.localSheetId;
		const hidden = entry.hidden;
		if (typeof name !== "string") {
			throw new XlsxError("invalid-input", `definedNames[${i}].name must be a string`);
		}
		const problem = definedNameProblem(name);
		if (problem !== undefined) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}].name ${definedNameProblemMessage(name, problem)}`,
			);
		}
		// `_xlnm._FilterDatabase` is a legal built-in name, but it is OWNED by a sheet's autoFilter (F10.2):
		// the writer synthesizes it from `SheetInput.autoFilter`, so a caller-supplied one would double it.
		if (isFilterDatabaseName(name)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}]: "${name}" is managed by SheetInput.autoFilter — set the sheet's autoFilter instead of a defined name`,
			);
		}
		if (typeof refersTo !== "string") {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must be a string`,
			);
		}
		if (refersTo.length === 0) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must not be empty`,
			);
		}
		if (refersTo.length > MAX_FORMULA_LEN) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo exceeds Excel's ${MAX_FORMULA_LEN}-character limit`,
			);
		}
		if (refersTo.startsWith("=")) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must be in stored form, without a leading "="`,
			);
		}
		if (!isXmlSafe(refersTo)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo contains a character not allowed in XML`,
			);
		}
		let scopeKey = "*"; // workbook-global scope
		if (localSheetId !== undefined) {
			if (
				typeof localSheetId !== "number" ||
				!Number.isInteger(localSheetId) ||
				localSheetId < 0 ||
				localSheetId >= sheetCount
			) {
				throw new XlsxError(
					"invalid-input",
					`definedNames[${i}] ("${name}"): localSheetId must be an integer index of an existing sheet (0..${sheetCount - 1})`,
				);
			}
			scopeKey = String(localSheetId);
		}
		if (hidden !== undefined && typeof hidden !== "boolean") {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): hidden must be a boolean`,
			);
		}
		const dupKey = `${scopeKey} ${name.toUpperCase()}`;
		if (seen.has(dupKey)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}]: duplicate name "${name}" in the same scope (names are case-insensitive)`,
			);
		}
		seen.add(dupKey);
		out.push({
			name,
			refersTo,
			...(localSheetId !== undefined ? { localSheetId } : {}),
			...(hidden === true ? { hidden: true } : {}),
		});
	}
	return out;
}

// Build a sheet-qualified ABSOLUTE A1 range like `'Data'!$A$1:$C$3` from a canonical range `A1:C3`. The
// sheet name is always single-quoted (a literal `'` doubles to `''`) — always valid, matches openpyxl,
// and sidesteps every unquoted-name edge (spaces, leading digits, cell-ref-shaped names). `ref` is
// pre-validated canonical, so parseCanonicalRange resolves; the fallback keeps the function total.
function qualifiedAbsoluteRange(sheetName: string, ref: string): string {
	const quoted = `'${sheetName.replace(/'/g, "''")}'`;
	const range = parseCanonicalRange(ref);
	if (range === undefined) return `${quoted}!${ref}`; // unreachable: ref is pre-validated canonical
	const abs = (c: CellRef): string => `$${indexToColumn(c.col)}$${c.row}`;
	const cells = ref.includes(":") ? `${abs(range.from)}:${abs(range.to)}` : abs(range.from);
	return `${quoted}!${cells}`;
}

/**
 * Synthesize the hidden, sheet-scoped `_xlnm._FilterDatabase` defined name a sheet's autoFilter needs
 * (F10.2). Excel records a filter range BOTH as `<autoFilter>` and this name; openjsxl owns the name so
 * a filter is represented once (the reader strips it, the writer re-creates it). `refersTo` is the range
 * sheet-qualified and absolute; `localSheetId` scopes it to its sheet. `ref` is the canonical range the
 * sheet writer already validated.
 */
export function filterDatabaseName(
	sheetName: string,
	sheetIndex: number,
	ref: string,
): DefinedName {
	return {
		name: FILTER_DATABASE_NAME,
		refersTo: qualifiedAbsoluteRange(sheetName, ref),
		localSheetId: sheetIndex,
		hidden: true,
	};
}

/** `xl/_rels/workbook.xml.rels` — worksheet targets plus styles.xml/theme1.xml when present. */
export function workbookRelsXml(
	sheetCount: number,
	needStyles: boolean,
	needTheme: boolean,
): string {
	const relItems = [
		...Array.from(
			{ length: sheetCount },
			(_, i) =>
				`<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
		),
		...(needStyles
			? [
					`<Relationship Id="rId${sheetCount + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>`,
				]
			: []),
		...(needTheme
			? [
					`<Relationship Id="rId${sheetCount + (needStyles ? 2 : 1)}" Type="${NS_REL}/theme" Target="theme/theme1.xml"/>`,
				]
			: []),
	].join("");
	return `${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${relItems}</Relationships>`;
}

/** A per-sheet `_rels` part (F4.6 hyperlinks, F5.2 comments + vmlDrawing). */
export function sheetRelsXml(rels: readonly SheetRel[]): string {
	const items = rels
		.map(
			(rel, j) =>
				`<Relationship Id="rId${j + 1}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"${rel.external ? ' TargetMode="External"' : ""}/>`,
		)
		.join("");
	return `${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${items}</Relationships>`;
}

/**
 * The per-sheet SIDE parts — the rels part and the comments + VML drawing pair — as `(name, xml)`
 * pairs, in the exact order both writers emit them (rels, then comments, then VML). This is the ONE
 * place that owns these OPC part-name conventions, so the buffered and streaming writers can't drift
 * on them (and the next per-sheet part family — drawings in M6 — has a single home to grow into).
 * `sheetIndex` is 0-based; parts are numbered `sheetIndex + 1`. The worksheet BODY part is emitted by
 * each writer itself (a string vs a chunk stream), so it stays out of here.
 */
export function sheetSideParts(
	sheetIndex: number,
	side: SheetSideParts,
): { name: string; xml: string }[] {
	const n = sheetIndex + 1;
	const parts: { name: string; xml: string }[] = [];
	if (side.rels.length > 0) {
		parts.push({
			name: `xl/worksheets/_rels/sheet${n}.xml.rels`,
			xml: sheetRelsXml(side.rels),
		});
	}
	if (side.commentsXml !== undefined) {
		parts.push({ name: `xl/comments${n}.xml`, xml: side.commentsXml });
	}
	if (side.vmlXml !== undefined) {
		parts.push({ name: `xl/drawings/vmlDrawing${n}.vml`, xml: side.vmlXml });
	}
	if (side.drawingXml !== undefined) {
		parts.push({ name: `xl/drawings/drawing${n}.xml`, xml: side.drawingXml });
	}
	if (side.drawingRelsXml !== undefined) {
		parts.push({
			name: `xl/drawings/_rels/drawing${n}.xml.rels`,
			xml: side.drawingRelsXml,
		});
	}
	// Table parts are numbered workbook-globally (not by sheet), so each carries its own number.
	if (side.tables !== undefined) {
		for (const table of side.tables) {
			parts.push({ name: `xl/tables/table${table.number}.xml`, xml: table.xml });
		}
	}
	return parts;
}

/**
 * The theme part to emit (F5.3): a caller-carried custom theme when present (validated non-empty and
 * XML-safe, read once), else the built-in Office theme. `carriedTheme` is read once by the caller and
 * passed here.
 */
export function themeToEmit(carriedTheme: unknown): string {
	if (carriedTheme === undefined) return DEFAULT_THEME_XML;
	if (typeof carriedTheme !== "string" || carriedTheme.length === 0) {
		throw new XlsxError("invalid-input", "themeXml must be a non-empty string");
	}
	if (!isXmlSafe(carriedTheme)) {
		throw new XlsxError(
			"invalid-input",
			"themeXml contains a character not allowed in XML (a control character or lone surrogate)",
		);
	}
	return carriedTheme;
}
