import { XlsxError } from "../errors";
import type { SheetState } from "../types";
import type { SheetRel } from "./sheet";
import { DEFAULT_THEME_XML } from "./theme";
import { escapeAttr, isXmlSafe } from "./xml";

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

/** `[Content_Types].xml` — one Override per part not covered by the rels/xml defaults. */
export function contentTypesXml(
	sheetCount: number,
	needStyles: boolean,
	needTheme: boolean,
	commentSheets: readonly number[],
	needVml: boolean,
): string {
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
	].join("");
	const vmlDefault = needVml
		? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
		: "";
	return `${XML_DECL}\n<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="${CT_RELS}"/><Default Extension="xml" ContentType="application/xml"/>${vmlDefault}${overrides}</Types>`;
}

/** `_rels/.rels` — the package's single relationship: officeDocument → the workbook. */
export function packageRelsXml(): string {
	return `${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
}

/**
 * `xl/workbook.xml` — the sheet list, with visibility state and the active-tab fix-up. `names` are the
 * VALIDATED names from {@link validateSheetMeta} (not re-read from the caller), so a getter can't slip
 * a different name in between validation and emission.
 */
export function workbookXml(
	names: readonly string[],
	states: readonly SheetState[],
	date1904: boolean,
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
	return `${XML_DECL}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${workbookPr}${bookViews}<sheets>${sheetsXml}</sheets></workbook>`;
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
