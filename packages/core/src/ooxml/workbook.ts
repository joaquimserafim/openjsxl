import type { SheetState } from "../types";
import { localName, relationshipId } from "../utils";
import { tokenize } from "../xml";

// xl/workbook.xml lists the workbook's sheets in tab order. Each <sheet> gives a display
// name and an r:id — NOT a filename — that points into workbook.xml.rels to locate the
// actual worksheet part. (sheetId is an internal key, not a file mapping.) It also carries
// the workbook-wide date system flag (<workbookPr date1904>).

export interface WorkbookSheet {
	/** Sheet name as shown on the tab. */
	readonly name: string;
	/** Relationship id (r:id) resolving to the worksheet part via workbook.xml.rels. */
	readonly rid: string;
	/** false for hidden or very-hidden sheets (the `state` attribute). */
	readonly visible: boolean;
	/** The tab's visibility state. Absent/unrecognized values read as "visible" (spec default). */
	readonly state: SheetState;
}

/**
 * A workbook-level defined (named) range or constant, from `<definedNames>`. `refersTo` is the raw
 * formula text (`'Sheet1'!$A$1:$B$2`, `42`, `Sheet1!$A:$A`). `localSheetId` (0-based sheet index)
 * marks a sheet-scoped name; absent means workbook-global. Built-in print ranges etc. keep their
 * `_xlnm.*` names. The evaluator resolves constants and simple ranges; anything else is `#NAME?` on
 * use (F8.2).
 */
export interface DefinedName {
	readonly name: string;
	readonly refersTo: string;
	readonly localSheetId?: number;
	readonly hidden?: boolean;
}

export interface WorkbookMeta {
	readonly sheets: readonly WorkbookSheet[];
	/** The 1904 date system flag (`<workbookPr date1904>`); selects the date serial epoch. */
	readonly date1904: boolean;
	/** Defined (named) ranges/constants, in document order. Empty when the workbook declares none. */
	readonly definedNames: readonly DefinedName[];
}

export function parseWorkbook(xml: string): WorkbookMeta {
	const sheets: WorkbookSheet[] = [];
	const definedNames: DefinedName[] = [];
	let date1904 = false;
	// A `<definedName>` carries its refersTo formula as TEXT content, so we accumulate across text
	// tokens between the open and close (mirroring parseFormulas' `<f>` handling).
	let dnName: string | undefined;
	let dnLocalSheetId: number | undefined;
	let dnHidden = false;
	let dnText = "";
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const tag = localName(token.name);
			if (tag === "workbookPr") {
				const flag = token.attrs.date1904;
				if (flag === "1" || flag === "true") date1904 = true;
			} else if (tag === "sheet") {
				const name = token.attrs.name;
				const rid = relationshipId(token.attrs);
				if (name === undefined || rid === undefined) continue;
				const raw = token.attrs.state;
				// Only the two hiding values are honoured; anything else (absent, "visible", garbage)
				// is the default. This keeps `visible` and `state` mechanically consistent.
				const state: SheetState =
					raw === "hidden" || raw === "veryHidden" ? raw : "visible";
				sheets.push({ name, rid, visible: state === "visible", state });
			} else if (tag === "definedName" && !token.selfClosing) {
				dnName = token.attrs.name;
				const rawId = token.attrs.localSheetId;
				const id = rawId !== undefined ? Number.parseInt(rawId, 10) : Number.NaN;
				dnLocalSheetId = Number.isInteger(id) && id >= 0 ? id : undefined;
				dnHidden = token.attrs.hidden === "1" || token.attrs.hidden === "true";
				dnText = "";
			}
		} else if (token.kind === "text") {
			if (dnName !== undefined) dnText += token.value;
		} else if (localName(token.name) === "definedName") {
			if (dnName !== undefined) {
				definedNames.push({
					name: dnName,
					refersTo: dnText,
					...(dnLocalSheetId !== undefined ? { localSheetId: dnLocalSheetId } : {}),
					...(dnHidden ? { hidden: true } : {}),
				});
				dnName = undefined;
			}
		}
	}
	return { sheets, date1904, definedNames };
}
