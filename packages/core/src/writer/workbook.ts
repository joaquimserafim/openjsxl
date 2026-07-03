import { XlsxError } from "../errors"
import type { SheetState } from "../types"
import { worksheetXml } from "./sheet"
import { createStyleRegistry } from "./styles"
import { DEFAULT_THEME_XML } from "./theme"
import type { WorkbookInput, WriteOptions } from "./types"
import { escapeAttr, isXmlSafe } from "./xml"
import { writeZip, type ZipInput } from "./zip"

// writeXlsx — the public workbook writer (F3.2, styled cells F4.2). Given a workbook described as
// plain data, emit the minimal valid OPC part set and pack it with writeZip. The output re-reads
// through openXlsx with the same values, types, sheet order, and styles, and opens in
// Excel/LibreOffice without a "repair" prompt.
//
// Minimal part set (only what a valid spreadsheet requires):
//   [Content_Types].xml            — content-type map for every part
//   _rels/.rels                    — package → xl/workbook.xml
//   xl/workbook.xml                — sheet list (names, ids, r:id links)
//   xl/_rels/workbook.xml.rels     — workbook → each worksheet (+ styles/theme when present)
//   xl/worksheets/sheetN.xml       — one per sheet
//   xl/styles.xml                  — ONLY when some cell needs a non-default format (a date, or
//                                    any F4.2 style); built by the shared StyleRegistry
//   xl/theme/theme1.xml            — ONLY when a written style uses a theme color; Excel resolves
//                                    {theme, tint} indexes against this part
//
// Strings are written inline (t="inlineStr"), so there is no sharedStrings.xml to buffer — the
// "value extractor" trade-off: a touch more bytes on disk for a single streaming pass and no table.

const encoder = new TextEncoder()
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const CT_BASE = "application/vnd.openxmlformats-officedocument.spreadsheetml"
const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"

// Excel's sheet-name rules: 1–31 chars, none of \ / ? * [ ] :, and unique ignoring case. A file that
// breaks these opens with a "repair" prompt, so we refuse to write it.
const MAX_SHEET_NAME = 31
const FORBIDDEN_SHEET_NAME = /[\\/?*[\]:]/

// Validate the workbook and resolve each sheet's visibility state in ONE read of the
// caller-supplied `state` property. Returning the resolved states (rather than re-reading
// `sheet.state` during emission) closes a TOCTOU gap: a getter/Proxy that answers "visible" here
// and "hidden" later could otherwise bypass the all-hidden guard or inject markup into the
// attribute. Downstream code emits from this array, never from `sheet.state`, so the value is
// always one of the three validated literals.
function validate(workbook: WorkbookInput): SheetState[] {
	const sheets = workbook?.sheets
	if (!Array.isArray(sheets) || sheets.length === 0) {
		throw new XlsxError("invalid-input", "a workbook needs at least one sheet")
	}
	const seen = new Set<string>()
	const states: SheetState[] = []
	let anyVisible = false
	for (const sheet of sheets) {
		const name = sheet?.name
		if (typeof name !== "string" || name.length === 0) {
			throw new XlsxError("invalid-input", "a sheet name must be a non-empty string")
		}
		if (name.length > MAX_SHEET_NAME) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" exceeds ${MAX_SHEET_NAME} characters`,
			)
		}
		if (FORBIDDEN_SHEET_NAME.test(name)) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" contains a forbidden character (\\ / ? * [ ] :)`,
			)
		}
		// A control character or lone surrogate in the name would corrupt the workbook.xml attribute
		// the same way it would a cell — the whole file would then fail to open.
		if (!isXmlSafe(name)) {
			throw new XlsxError(
				"invalid-input",
				`sheet name "${name}" contains a character not allowed in XML`,
			)
		}
		// Excel treats sheet names case-insensitively, so "Data" and "data" collide.
		const key = name.toLowerCase()
		if (seen.has(key)) {
			throw new XlsxError(
				"invalid-input",
				`duplicate sheet name "${name}" (case-insensitive)`,
			)
		}
		seen.add(key)
		if (!Array.isArray(sheet.rows)) {
			throw new XlsxError("invalid-input", `sheet "${name}": rows must be an array`)
		}
		// Tab visibility (F4.6). Read the caller's property ONCE (getters/Proxies can vary between
		// reads); absent means visible; only the three spec values are accepted.
		const state = sheet.state ?? "visible"
		if (state !== "visible" && state !== "hidden" && state !== "veryHidden") {
			throw new XlsxError(
				"invalid-input",
				`sheet "${name}": state must be "visible", "hidden", or "veryHidden"`,
			)
		}
		states.push(state)
		if (state === "visible") anyVisible = true
	}
	// Excel refuses a workbook whose every sheet is hidden — there would be nothing to show.
	if (!anyVisible) {
		throw new XlsxError("invalid-input", "at least one sheet must be visible")
	}
	return states
}

/**
 * Serialize a workbook to `.xlsx` bytes. Async because compression runs on the platform's
 * CompressionStream. Throws {@link XlsxError} with code `invalid-input` for anything that can't be
 * represented — no sheets, a bad sheet name, or a cell value that isn't a string, finite number,
 * boolean, Date, or null.
 */
export async function writeXlsx(
	workbook: WorkbookInput,
	options?: WriteOptions,
): Promise<Uint8Array> {
	const states = validate(workbook)
	const date1904 = options?.date1904 === true
	const sheets = workbook.sheets

	// Render every worksheet up front — this is also where invalid cell values and styles surface —
	// interning every style into one shared registry. What the registry saw decides whether
	// styles.xml (any non-default format) and theme1.xml (any theme color) are emitted at all.
	const styles = createStyleRegistry()
	const worksheets = sheets.map((sheet, i) => worksheetXml(sheet, i, date1904, styles))
	const needStyles = styles.needed()
	const needTheme = styles.usesTheme()

	// Comments (F5.2) each ride on a legacy VML drawing part; a workbook with any comment gains the
	// `vml` Default content type, and every comments part gets an Override. Sheets are named by index
	// so a comments/VML part never collides with another sheet's.
	const commentSheets = worksheets.flatMap((w, i) => (w.commentsXml !== undefined ? [i] : []))
	const needVml = commentSheets.length > 0

	const parts: ZipInput[] = []
	const add = (name: string, xml: string): void => {
		parts.push({ name, data: encoder.encode(xml) })
	}

	// [Content_Types].xml — one Override per part that isn't covered by the rels/xml defaults.
	const overrides = [
		`<Override PartName="/xl/workbook.xml" ContentType="${CT_BASE}.sheet.main+xml"/>`,
		...sheets.map(
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
	].join("")
	// The `vml` Default covers every vmlDrawing part (F5.2); emitted only when a comment exists, so a
	// comment-free workbook keeps its exact pre-F5.2 [Content_Types].xml bytes.
	const vmlDefault = needVml
		? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
		: ""
	add(
		"[Content_Types].xml",
		`${XML_DECL}\n<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="${CT_RELS}"/><Default Extension="xml" ContentType="application/xml"/>${vmlDefault}${overrides}</Types>`,
	)

	// _rels/.rels — the package's single relationship: officeDocument → the workbook.
	add(
		"_rels/.rels",
		`${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	)

	// xl/workbook.xml — the sheet list. rId(i+1) links each <sheet> to its worksheet rel below.
	const workbookPr = date1904 ? '<workbookPr date1904="1"/>' : ""
	// The active tab defaults to index 0; when the FIRST sheet is hidden that default would point
	// at a tab the user can't see, so aim it at the first visible sheet instead (validate()
	// guarantees one exists) — exactly what openpyxl does. All-visible workbooks emit no
	// <bookViews> and keep their exact pre-F4.6 bytes.
	const firstVisible = states.indexOf("visible")
	const bookViews =
		firstVisible > 0 ? `<bookViews><workbookView activeTab="${firstVisible}"/></bookViews>` : ""
	const sheetsXml = sheets
		.map((sheet, i) => {
			// Emit from the validated `states` array, never from `sheet.state` (single-read: a
			// caller's getter can't vary the value between validation and here). "visible" is the
			// spec default and stays implicit, so an all-visible workbook keeps its exact pre-F4.6
			// bytes; the value is always one of the three literals, so no escaping is needed.
			const state = states[i] === "visible" ? "" : ` state="${states[i]}"`
			return `<sheet name="${escapeAttr(sheet.name)}" sheetId="${i + 1}"${state} r:id="rId${i + 1}"/>`
		})
		.join("")
	add(
		"xl/workbook.xml",
		`${XML_DECL}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${workbookPr}${bookViews}<sheets>${sheetsXml}</sheets></workbook>`,
	)

	// xl/_rels/workbook.xml.rels — worksheet targets (matching the r:ids above) plus styles.xml and
	// theme1.xml when present. The styles rel takes the id after the last sheet; theme follows it.
	const relItems = [
		...sheets.map(
			(_, i) =>
				`<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
		),
		...(needStyles
			? [
					`<Relationship Id="rId${sheets.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>`,
				]
			: []),
		...(needTheme
			? [
					`<Relationship Id="rId${sheets.length + (needStyles ? 2 : 1)}" Type="${NS_REL}/theme" Target="theme/theme1.xml"/>`,
				]
			: []),
	].join("")
	add(
		"xl/_rels/workbook.xml.rels",
		`${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${relItems}</Relationships>`,
	)

	if (needStyles) add("xl/styles.xml", styles.stylesXml())
	if (needTheme) add("xl/theme/theme1.xml", DEFAULT_THEME_XML)

	worksheets.forEach((w, i) => {
		add(`xl/worksheets/sheet${i + 1}.xml`, w.xml)
		// The per-sheet rels part (F4.6 hyperlinks, F5.2 comments + vmlDrawing): one <Relationship>
		// per entry, ids matching the r:ids in the sheet XML. Hyperlink targets carry
		// TargetMode="External"; internal parts don't. Covered by the `Default Extension="rels"`
		// content type — no Override needed. A hyperlinks-only sheet reproduces its pre-F5.2 bytes.
		if (w.rels.length > 0) {
			const rels = w.rels
				.map(
					(rel, j) =>
						`<Relationship Id="rId${j + 1}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"${rel.external ? ' TargetMode="External"' : ""}/>`,
				)
				.join("")
			add(
				`xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
				`${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${rels}</Relationships>`,
			)
		}
		// Comments (F5.2): the comments part and its paired VML legacy drawing, named by sheet index
		// to match the rel targets built in worksheetXml.
		if (w.commentsXml !== undefined) add(`xl/comments${i + 1}.xml`, w.commentsXml)
		if (w.vmlXml !== undefined) add(`xl/drawings/vmlDrawing${i + 1}.vml`, w.vmlXml)
	})

	return writeZip(parts)
}
