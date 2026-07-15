import type { DxfStyle, TableColumn, TableInfo, TableStyleInfo } from "../types";
import { localName } from "../utils";
import { tokenize } from "../xml";
import { parseRef } from "./a1";
import { normalizeName } from "./name";
import { decodeXstring } from "./xstring";

// Table part parser (F9.1). Reads `xl/tables/tableN.xml` into the shared {@link TableInfo} model. The
// reader is TOLERANT: absent attributes degrade to Excel's defaults (header on, totals off), and a
// part with no usable name or ref yields `undefined` (the caller drops it) rather than throwing.

// A table display name obeys the shared defined-name-style identifier grammar (decision 8). F10.1
// extracted that core to ooxml/name.ts (a table name and a workbook defined name are the same grammar,
// single-sourced so they can never drift). These re-exports keep the historical table-name spelling
// stable for importers and tests.
export {
	MAX_NAME_LEN as MAX_TABLE_NAME_LEN,
	type NameProblem as TableNameProblem,
	nameProblem as tableNameProblem,
	normalizeName as normalizeTableName,
} from "./name";

/**
 * The width/height (in cells) of an A1 range like `"A1:C5"` — or a single cell `"A1"` — or `undefined`
 * when the ref doesn't parse. Used to reconcile a table's column count and totals row to the strict
 * writer's rules (F9.5). `parseRef` throws on a malformed token, so it is shielded here.
 */
function rangeExtent(ref: string): { width: number; height: number } | undefined {
	const parts = ref.split(":");
	if (parts.length > 2) return undefined;
	const a = parts[0];
	if (a === undefined) return undefined;
	const b = parts[1] ?? a; // a single cell "A1" has no ":" — both ends are the same
	try {
		const start = parseRef(a);
		const end = parseRef(b);
		return {
			width: Math.abs(end.col - start.col) + 1,
			height: Math.abs(end.row - start.row) + 1,
		};
	} catch {
		return undefined;
	}
}

// OOXML booleans are `1`/`0` (or `true`/`false`); anything else — or an absent attribute — is undefined.
function bool01(v: string | undefined): boolean | undefined {
	if (v === "1" || v === "true") return true;
	if (v === "0" || v === "false") return false;
	return undefined;
}

// A mutable column builder — the formula fields arrive as CHILD element text, filled in between the
// column's open and close tokens.
interface ColumnBuilder {
	name: string;
	totalsRowLabel?: string;
	totalsRowFunction?: string;
	totalsRowFormula?: string;
	calculatedColumnFormula?: string;
	headerRowDxfId?: number;
	dataDxfId?: number;
	totalsRowDxfId?: number;
}

// A non-negative integer attribute (a dxf index), or undefined when absent/malformed. Gated to
// CANONICAL decimal digits first — `Number("")`/`Number(" 0 ")`/`Number("1e2")`/`Number("0x1")` all
// coerce to valid integers, which would fabricate a phantom highlight from an empty/odd attribute.
function intId(v: string | undefined): number | undefined {
	if (v === undefined || !/^[0-9]+$/.test(v)) return undefined;
	const n = Number(v);
	return Number.isInteger(n) ? n : undefined;
}

// Resolve a dxf index to its inline style (F9.3). Out-of-range/absent → undefined (no highlight).
function resolveDxf(
	id: number | undefined,
	dxfs: readonly DxfStyle[] | undefined,
): DxfStyle | undefined {
	return id !== undefined ? dxfs?.[id] : undefined;
}

/**
 * Parse a table part into a {@link TableInfo}, or `undefined` when it lacks a usable name/ref. `dxfs`
 * (the workbook `<dxfs>` table, F9.3) resolves the table's/columns' `*DxfId` highlight indexes to
 * inline {@link DxfStyle}s; absent when the workbook has no differential styles.
 */
export function parseTable(xml: string, dxfs?: readonly DxfStyle[]): TableInfo | undefined {
	let name: string | undefined;
	let ref: string | undefined;
	let headerRowCount: string | undefined;
	let totalsRowCount: string | undefined;
	let style: TableStyleInfo | undefined;
	// Table-level highlight dxf indexes (resolved at the end).
	let tableHeaderDxfId: number | undefined;
	let tableDataDxfId: number | undefined;
	let tableTotalsDxfId: number | undefined;
	const columns: ColumnBuilder[] = [];
	let column: ColumnBuilder | undefined; // the tableColumn currently open
	let textTarget: "totals" | "calc" | undefined; // which formula child is accumulating text
	let text = "";

	for (const tok of tokenize(xml)) {
		if (tok.kind === "open") {
			switch (localName(tok.name)) {
				case "table": {
					// displayName/name are ST_Xstring (F9.6) — Excel and openpyxl decode them, so
					// a stored `_x005F_x0041_` IS the name `_x0041_`; the writer encodes them back.
					const rawName = tok.attrs.displayName ?? tok.attrs.name;
					name = rawName !== undefined ? decodeXstring(rawName) : undefined;
					ref = tok.attrs.ref;
					headerRowCount = tok.attrs.headerRowCount;
					totalsRowCount = tok.attrs.totalsRowCount;
					tableHeaderDxfId = intId(tok.attrs.headerRowDxfId);
					tableDataDxfId = intId(tok.attrs.dataDxfId);
					tableTotalsDxfId = intId(tok.attrs.totalsRowDxfId);
					break;
				}
				case "tableColumn": {
					const cn = tok.attrs.name;
					if (cn !== undefined) {
						// @name and @totalsRowLabel are ST_Xstring, like the header cells they
						// mirror — decode so the column name equals its decoded header cell (F9.6).
						const builder: ColumnBuilder = { name: decodeXstring(cn) };
						if (tok.attrs.totalsRowLabel !== undefined)
							builder.totalsRowLabel = decodeXstring(tok.attrs.totalsRowLabel);
						if (tok.attrs.totalsRowFunction !== undefined)
							builder.totalsRowFunction = tok.attrs.totalsRowFunction;
						const hId = intId(tok.attrs.headerRowDxfId);
						if (hId !== undefined) builder.headerRowDxfId = hId;
						const dId = intId(tok.attrs.dataDxfId);
						if (dId !== undefined) builder.dataDxfId = dId;
						const tId = intId(tok.attrs.totalsRowDxfId);
						if (tId !== undefined) builder.totalsRowDxfId = tId;
						if (tok.selfClosing) columns.push(builder);
						else column = builder;
					}
					break;
				}
				case "totalsRowFormula":
					textTarget = tok.selfClosing ? undefined : "totals";
					text = "";
					break;
				case "calculatedColumnFormula":
					textTarget = tok.selfClosing ? undefined : "calc";
					text = "";
					break;
				case "tableStyleInfo":
					style = readStyleInfo(tok.attrs);
					break;
			}
		} else if (tok.kind === "text") {
			if (textTarget !== undefined) text += tok.value;
		} else {
			switch (localName(tok.name)) {
				case "tableColumn":
					if (column !== undefined) {
						columns.push(column);
						column = undefined;
					}
					break;
				case "totalsRowFormula":
					if (column !== undefined && textTarget === "totals")
						column.totalsRowFormula = text;
					textTarget = undefined;
					break;
				case "calculatedColumnFormula":
					if (column !== undefined && textTarget === "calc")
						column.calculatedColumnFormula = text;
					textTarget = undefined;
					break;
			}
		}
	}

	if (name === undefined || ref === undefined) return undefined;
	const legalName = normalizeName(name); // F9.5: degrade an odd name to a writer-legal identifier
	// Resolve each column's highlight dxf indexes to inline styles (F9.3 retrofit).
	const outColumns: TableColumn[] = columns.map((c) => {
		const col: {
			name: string;
			totalsRowLabel?: string;
			totalsRowFunction?: string;
			totalsRowFormula?: string;
			calculatedColumnFormula?: string;
			headerRowStyle?: DxfStyle;
			dataStyle?: DxfStyle;
			totalsRowStyle?: DxfStyle;
		} = { name: c.name };
		if (c.totalsRowLabel !== undefined) col.totalsRowLabel = c.totalsRowLabel;
		if (c.totalsRowFunction !== undefined) col.totalsRowFunction = c.totalsRowFunction;
		if (c.totalsRowFormula !== undefined) col.totalsRowFormula = c.totalsRowFormula;
		if (c.calculatedColumnFormula !== undefined)
			col.calculatedColumnFormula = c.calculatedColumnFormula;
		const h = resolveDxf(c.headerRowDxfId, dxfs);
		if (h !== undefined) col.headerRowStyle = h;
		const d = resolveDxf(c.dataDxfId, dxfs);
		if (d !== undefined) col.dataStyle = d;
		const t = resolveDxf(c.totalsRowDxfId, dxfs);
		if (t !== undefined) col.totalsRowStyle = t;
		return col;
	});
	const tHeader = resolveDxf(tableHeaderDxfId, dxfs);
	const tData = resolveDxf(tableDataDxfId, dxfs);
	const tTotals = resolveDxf(tableTotalsDxfId, dxfs);
	// F9.5: reconcile shape to the strict writer's rules. A NON-EMPTY `columns` must have one entry per
	// ref column — on a mismatch, clear it so the writer derives names from the header row. A totals row
	// needs a data row above it — a single-row ref can't carry one, so clear it.
	const extent = rangeExtent(ref);
	const reconciledColumns =
		extent !== undefined && outColumns.length > 0 && outColumns.length !== extent.width
			? []
			: outColumns;
	const wantsTotals = totalsRowCount !== undefined && totalsRowCount !== "0";
	const totalsRow = wantsTotals && !(extent !== undefined && extent.height === 1);
	const info: TableInfo = {
		name: legalName,
		ref,
		columns: reconciledColumns,
		headerRow: headerRowCount !== "0",
		totalsRow,
		...(style !== undefined ? { style } : {}),
		...(tHeader !== undefined ? { headerRowStyle: tHeader } : {}),
		...(tData !== undefined ? { dataStyle: tData } : {}),
		...(tTotals !== undefined ? { totalsRowStyle: tTotals } : {}),
	};
	return info;
}

function readStyleInfo(attrs: Readonly<Record<string, string>>): TableStyleInfo {
	const style: {
		name?: string;
		showFirstColumn?: boolean;
		showLastColumn?: boolean;
		showRowStripes?: boolean;
		showColumnStripes?: boolean;
	} = {};
	if (attrs.name !== undefined) style.name = attrs.name;
	const first = bool01(attrs.showFirstColumn);
	if (first !== undefined) style.showFirstColumn = first;
	const last = bool01(attrs.showLastColumn);
	if (last !== undefined) style.showLastColumn = last;
	const rows = bool01(attrs.showRowStripes);
	if (rows !== undefined) style.showRowStripes = rows;
	const cols = bool01(attrs.showColumnStripes);
	if (cols !== undefined) style.showColumnStripes = cols;
	return style;
}
