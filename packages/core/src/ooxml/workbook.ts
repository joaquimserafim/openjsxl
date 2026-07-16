import type { SheetState, WorkbookProtection } from "../types";
import { localName, relationshipId } from "../utils";
import { tokenize } from "../xml";
import { definedNameEmittable, isFilterDatabaseName } from "./name";
import { MAX_SPIN_COUNT } from "./styles";
import { decodeXstring } from "./xstring";

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
	/** Workbook `<workbookProtection>`, or `undefined` when none (F10.3). */
	readonly protection: WorkbookProtection | undefined;
}

// <workbookProtection>'s boolean attributes (F10.3), carried verbatim alongside any password material.
function readWorkbookProtection(
	attrs: Readonly<Record<string, string>>,
): WorkbookProtection | undefined {
	const bool = (v: string | undefined): boolean | undefined =>
		v === "1" || v === "true" ? true : v === "0" || v === "false" ? false : undefined;
	const out: {
		lockStructure?: boolean;
		lockWindows?: boolean;
		workbookPassword?: string;
		workbookAlgorithmName?: string;
		workbookHashValue?: string;
		workbookSaltValue?: string;
		workbookSpinCount?: number;
	} = {};
	const lockStructure = bool(attrs.lockStructure);
	if (lockStructure !== undefined) out.lockStructure = lockStructure;
	const lockWindows = bool(attrs.lockWindows);
	if (lockWindows !== undefined) out.lockWindows = lockWindows;
	if (attrs.workbookPassword !== undefined) out.workbookPassword = attrs.workbookPassword;
	if (attrs.workbookAlgorithmName !== undefined)
		out.workbookAlgorithmName = attrs.workbookAlgorithmName;
	if (attrs.workbookHashValue !== undefined) out.workbookHashValue = attrs.workbookHashValue;
	if (attrs.workbookSaltValue !== undefined) out.workbookSaltValue = attrs.workbookSaltValue;
	// xsd:unsignedInt — DROP an out-of-range count (shared bound with the writer's reject; see MAX_SPIN_COUNT).
	const spin = attrs.workbookSpinCount;
	if (spin !== undefined && /^[0-9]+$/.test(spin)) {
		const n = Number(spin);
		if (Number.isInteger(n) && n <= MAX_SPIN_COUNT) out.workbookSpinCount = n;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

export function parseWorkbook(xml: string): WorkbookMeta {
	const sheets: WorkbookSheet[] = [];
	const definedNames: DefinedName[] = [];
	let date1904 = false;
	let protection: WorkbookProtection | undefined;
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
			} else if (tag === "workbookProtection") {
				protection = readWorkbookProtection(token.attrs);
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
				// @name is ST_Xstring, the same schema type as a table displayName (F9.6) — Excel and
				// openpyxl decode it, so a stored `_x005F_x0041_` IS the name `_x0041_`. Decode here so
				// the model name equals its true value and re-encodes losslessly on write.
				dnName =
					token.attrs.name !== undefined ? decodeXstring(token.attrs.name) : undefined;
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
	// F10.1 decision 3: keep only names the strict writer could re-emit, so `Workbook.definedNames`
	// holds writer-legal entries only (the shared-model invariant — what the reader returns IS what
	// the writer accepts). A foreign producer's illegal name, an empty/oversized/XML-unsafe `refersTo`,
	// or a sheet-scope pointing past the sheet list is DROPPED — a named, tested degradation. A defined
	// name is dropped rather than normalized (unlike a table name): it is referenced by formulas, so
	// renaming it would silently break those links. Such an illegal name is also unreferenceable by any
	// valid formula, so dropping it never changes what the evaluator resolves.
	// Strip the reserved `_xlnm._FilterDatabase` (F10.2): it is internal bookkeeping owned by a sheet's
	// autoFilter, surfaced as `Worksheet.autoFilter` and re-synthesized on write — never a public defined
	// name (matches openpyxl, which reports `defined_names == []` for a filtered sheet). Then keep only
	// names the strict writer could re-emit (decision 3, shared-model invariant).
	const emittable = definedNames.filter(
		(dn) => !isFilterDatabaseName(dn.name) && definedNameEmittable(dn, sheets.length),
	);
	return { sheets, date1904, definedNames: emittable, protection };
}
