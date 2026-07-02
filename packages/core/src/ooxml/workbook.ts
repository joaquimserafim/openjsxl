import type { SheetState } from "../types"
import { localName, relationshipId } from "../utils"
import { tokenize } from "../xml"

// xl/workbook.xml lists the workbook's sheets in tab order. Each <sheet> gives a display
// name and an r:id — NOT a filename — that points into workbook.xml.rels to locate the
// actual worksheet part. (sheetId is an internal key, not a file mapping.) It also carries
// the workbook-wide date system flag (<workbookPr date1904>).

export interface WorkbookSheet {
	/** Sheet name as shown on the tab. */
	readonly name: string
	/** Relationship id (r:id) resolving to the worksheet part via workbook.xml.rels. */
	readonly rid: string
	/** false for hidden or very-hidden sheets (the `state` attribute). */
	readonly visible: boolean
	/** The tab's visibility state. Absent/unrecognized values read as "visible" (spec default). */
	readonly state: SheetState
}

export interface WorkbookMeta {
	readonly sheets: readonly WorkbookSheet[]
	/** The 1904 date system flag (`<workbookPr date1904>`); selects the date serial epoch. */
	readonly date1904: boolean
}

export function parseWorkbook(xml: string): WorkbookMeta {
	const sheets: WorkbookSheet[] = []
	let date1904 = false
	for (const token of tokenize(xml)) {
		if (token.kind !== "open") continue
		const tag = localName(token.name)
		if (tag === "workbookPr") {
			const flag = token.attrs.date1904
			if (flag === "1" || flag === "true") date1904 = true
		} else if (tag === "sheet") {
			const name = token.attrs.name
			const rid = relationshipId(token.attrs)
			if (name === undefined || rid === undefined) continue
			const raw = token.attrs.state
			// Only the two hiding values are honoured; anything else (absent, "visible", garbage)
			// is the default. This keeps `visible` and `state` mechanically consistent.
			const state: SheetState = raw === "hidden" || raw === "veryHidden" ? raw : "visible"
			sheets.push({ name, rid, visible: state === "visible", state })
		}
	}
	return { sheets, date1904 }
}
