import { localName } from '../utils'
import { tokenize } from '../xml'

// xl/workbook.xml lists the workbook's sheets in tab order. Each <sheet> gives a display
// name and an r:id — NOT a filename — that points into workbook.xml.rels to locate the
// actual worksheet part. (sheetId is an internal key, not a file mapping.) It also carries
// the workbook-wide date system flag (<workbookPr date1904>).

export interface WorkbookSheet {
	/** Sheet name as shown on the tab. */
	name: string
	/** Relationship id (r:id) resolving to the worksheet part via workbook.xml.rels. */
	rid: string
	/** false for hidden or very-hidden sheets (the `state` attribute). */
	visible: boolean
}

export interface WorkbookMeta {
	sheets: WorkbookSheet[]
	/** The 1904 date system flag (`<workbookPr date1904>`); selects the date serial epoch. */
	date1904: boolean
}

// The relationship id is conventionally `r:id`, but the `r` prefix is only bound by
// convention — fall back to any attribute whose local name is `id`.
function relationshipId(attrs: Record<string, string>): string | undefined {
	if (attrs['r:id'] !== undefined) return attrs['r:id']
	for (const key of Object.keys(attrs)) {
		if (localName(key) === 'id') return attrs[key]
	}
	return undefined
}

export function parseWorkbook(xml: string): WorkbookMeta {
	const sheets: WorkbookSheet[] = []
	let date1904 = false
	for (const token of tokenize(xml)) {
		if (token.kind !== 'open') continue
		const tag = localName(token.name)
		if (tag === 'workbookPr') {
			const flag = token.attrs.date1904
			if (flag === '1' || flag === 'true') date1904 = true
		} else if (tag === 'sheet') {
			const name = token.attrs.name
			const rid = relationshipId(token.attrs)
			if (name === undefined || rid === undefined) continue
			const state = token.attrs.state
			sheets.push({ name, rid, visible: state !== 'hidden' && state !== 'veryHidden' })
		}
	}
	return { sheets, date1904 }
}
