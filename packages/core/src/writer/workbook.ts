import { XlsxError } from '../errors'
import { worksheetXml } from './sheet'
import type { WorkbookInput, WriteOptions } from './types'
import { escapeAttr, isXmlSafe } from './xml'
import { writeZip, type ZipInput } from './zip'

// writeXlsx — the public workbook writer (F3.2). Given a workbook described as plain data, emit the
// minimal valid OPC part set and pack it with writeZip. The output re-reads through openXlsx with the
// same values, types, and sheet order, and opens in Excel/LibreOffice without a "repair" prompt.
//
// Minimal part set (only what a valid spreadsheet requires):
//   [Content_Types].xml            — content-type map for every part
//   _rels/.rels                    — package → xl/workbook.xml
//   xl/workbook.xml                — sheet list (names, ids, r:id links)
//   xl/_rels/workbook.xml.rels     — workbook → each worksheet (+ styles when present)
//   xl/worksheets/sheetN.xml       — one per sheet
//   xl/styles.xml                  — ONLY when some cell is a date (holds the date cellXf)
//
// Strings are written inline (t="inlineStr"), so there is no sharedStrings.xml to buffer — the
// "value extractor" trade-off: a touch more bytes on disk for a single streaming pass and no table.

const encoder = new TextEncoder()
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const CT_BASE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml'

// Excel's sheet-name rules: 1–31 chars, none of \ / ? * [ ] :, and unique ignoring case. A file that
// breaks these opens with a "repair" prompt, so we refuse to write it.
const MAX_SHEET_NAME = 31
const FORBIDDEN_SHEET_NAME = /[\\/?*[\]:]/

// styles.xml with exactly two cellXfs: index 0 = General, index 1 = the built-in date format
// (numFmtId 14). Date cells carry s="1"; the reader maps numFmtId 14 back to a date. Emitted only
// when a date is present, keeping a date-free workbook down to the bare parts.
const STYLES_XML = `${XML_DECL}
<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`

function validate(workbook: WorkbookInput): void {
	const sheets = workbook?.sheets
	if (!Array.isArray(sheets) || sheets.length === 0) {
		throw new XlsxError('invalid-input', 'a workbook needs at least one sheet')
	}
	const seen = new Set<string>()
	for (const sheet of sheets) {
		const name = sheet?.name
		if (typeof name !== 'string' || name.length === 0) {
			throw new XlsxError('invalid-input', 'a sheet name must be a non-empty string')
		}
		if (name.length > MAX_SHEET_NAME) {
			throw new XlsxError(
				'invalid-input',
				`sheet name "${name}" exceeds ${MAX_SHEET_NAME} characters`,
			)
		}
		if (FORBIDDEN_SHEET_NAME.test(name)) {
			throw new XlsxError(
				'invalid-input',
				`sheet name "${name}" contains a forbidden character (\\ / ? * [ ] :)`,
			)
		}
		// A control character or lone surrogate in the name would corrupt the workbook.xml attribute
		// the same way it would a cell — the whole file would then fail to open.
		if (!isXmlSafe(name)) {
			throw new XlsxError(
				'invalid-input',
				`sheet name "${name}" contains a character not allowed in XML`,
			)
		}
		// Excel treats sheet names case-insensitively, so "Data" and "data" collide.
		const key = name.toLowerCase()
		if (seen.has(key)) {
			throw new XlsxError(
				'invalid-input',
				`duplicate sheet name "${name}" (case-insensitive)`,
			)
		}
		seen.add(key)
		if (!Array.isArray(sheet.rows)) {
			throw new XlsxError('invalid-input', `sheet "${name}": rows must be an array`)
		}
	}
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
	validate(workbook)
	const date1904 = options?.date1904 === true
	const sheets = workbook.sheets

	// Render every worksheet up front — this is also where invalid cell values surface — and learn
	// whether any sheet needs the date style (which decides if styles.xml is emitted at all).
	const worksheets = sheets.map((sheet) => worksheetXml(sheet.rows, date1904))
	const anyDate = worksheets.some((w) => w.usesDate)

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
		...(anyDate
			? [`<Override PartName="/xl/styles.xml" ContentType="${CT_BASE}.styles+xml"/>`]
			: []),
	].join('')
	add(
		'[Content_Types].xml',
		`${XML_DECL}\n<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="${CT_RELS}"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
	)

	// _rels/.rels — the package's single relationship: officeDocument → the workbook.
	add(
		'_rels/.rels',
		`${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	)

	// xl/workbook.xml — the sheet list. rId(i+1) links each <sheet> to its worksheet rel below.
	const workbookPr = date1904 ? '<workbookPr date1904="1"/>' : ''
	const sheetsXml = sheets
		.map(
			(sheet, i) =>
				`<sheet name="${escapeAttr(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
		)
		.join('')
	add(
		'xl/workbook.xml',
		`${XML_DECL}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${workbookPr}<sheets>${sheetsXml}</sheets></workbook>`,
	)

	// xl/_rels/workbook.xml.rels — worksheet targets (matching the r:ids above) plus styles.xml when
	// present. The styles rel takes the id after the last sheet.
	const relItems = [
		...sheets.map(
			(_, i) =>
				`<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
		),
		...(anyDate
			? [
					`<Relationship Id="rId${sheets.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>`,
				]
			: []),
	].join('')
	add(
		'xl/_rels/workbook.xml.rels',
		`${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${relItems}</Relationships>`,
	)

	if (anyDate) add('xl/styles.xml', STYLES_XML)

	worksheets.forEach((w, i) => {
		add(`xl/worksheets/sheet${i + 1}.xml`, w.xml)
	})

	return writeZip(parts)
}
