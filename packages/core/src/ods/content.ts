import { formatRef, MAX_COL, MAX_ROW } from "../ooxml";
import type { Cell, Hyperlink, SheetState } from "../types";
import { tokenize } from "../xml";

// Parse an ODF spreadsheet's content.xml into per-sheet typed cell tables. ODF is a different
// vocabulary over the SAME SAX tokenizer the xlsx reader uses: sheets are <table:table> inside
// <office:spreadsheet>, and cell TYPES are EXPLICIT (office:value-type + a typed value attribute),
// so there is no style-driven date detection — a date cell arrives already typed. Namespace
// prefixes (office:/table:/text:/xlink:) are the conventional, stable ODF ones and are matched
// literally, exactly as the xlsx reader matches r:id.
//
// The adversarial shape unique to ODS is the REPEAT: a cell/row may declare
// table:number-columns-repeated / table:number-rows-repeated up to the 2^20 grid edge (LibreOffice
// pads every sheet's tail this way). We never materialize an EMPTY repeat — an empty cell/row emits
// nothing and only advances the position cursor (saturated at the grid bounds) — so a
// repeat-to-the-edge tail costs O(1). Non-empty repeats DO materialize (a value filled across cells
// is real content), clamped to the grid and to a per-sheet cell ceiling, so total materialization is
// bounded by content or the cap, never by an attacker-chosen repeat count.

/** A parsed ODF sheet: the shared cell/merge/link model plus its tab identity. */
export interface OdsSheet {
	readonly name: string;
	readonly visible: boolean;
	readonly state: SheetState;
	readonly cells: Map<string, Cell>;
	readonly merges: readonly string[];
	readonly hyperlinks: readonly Hyperlink[];
	readonly dimension: string | undefined;
}

// Cap accumulated text per cell so a `<text:s text:c="1e9"/>` space run can't balloon memory.
const MAX_CELL_TEXT = 32_768;
// Cap materialized cells per sheet: legit sheets sit far below this; a non-empty repeat bomb
// degrades (cells past the cap are dropped) instead of expanding to the full 2^34-cell grid.
const MAX_ODS_CELLS = 2_000_000;

// A non-empty cell decoded but not yet placed at a final A1 ref (row placement is deferred because a
// row can itself repeat). `type`/`value` are a valid Cell pairing by construction.
interface DecodedCell {
	readonly type: Cell["type"];
	readonly value: string | number | boolean | Date | null;
}

interface PendingCell {
	/** 1-based column of the first instance. */
	readonly col: number;
	readonly decoded: DecodedCell;
	readonly href: string | undefined;
	readonly spanCols: number;
	readonly spanRows: number;
}

function intAttr(value: string | undefined, fallback: number): number {
	if (value === undefined) return fallback;
	const n = Number.parseInt(value, 10);
	return Number.isFinite(n) && n >= 1 ? n : fallback;
}

// ODF `office:date-value` is an ISO date or date-time; map it to the SAME UTC instant the xlsx path
// produces (every Date in this codebase is UTC-anchored wall-clock, no timezone — like Excel serials).
// A trailing timezone offset, if any, is ignored (naive wall-clock). Returns undefined on a bad value.
function parseOdsDate(value: string | undefined): Date | undefined {
	if (value === undefined) return undefined;
	const m = /^(-?\d{1,6})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?)?/.exec(
		value,
	);
	if (m === null) return undefined;
	const year = Number(m[1]);
	const month = Number(m[2]);
	const day = Number(m[3]);
	const hour = m[4] === undefined ? 0 : Number(m[4]);
	const min = m[5] === undefined ? 0 : Number(m[5]);
	const sec = m[6] === undefined ? 0 : Number(m[6]);
	const ms = m[7] === undefined ? 0 : Math.round(Number(`0.${m[7]}`) * 1000);
	const t = Date.UTC(year, month - 1, day, hour, min, sec, ms);
	if (Number.isNaN(t)) return undefined;
	const d = new Date(t);
	// `Date.UTC` applies the ECMAScript legacy two-digit-year rule to a year ARGUMENT in 0–99,
	// mapping it to 1900+year — so a spec-valid ODF year like `0050` would read as 1950. Undo it:
	// `setUTCFullYear` takes the literal year (no remap), preserving month/day/time already set.
	if (year >= 0 && year <= 99) d.setUTCFullYear(year);
	return d;
}

// ODF `office:time-value` is an ISO 8601 duration (PnDTnHnMnS). Time-of-day is stored as a duration
// from midnight; we return it as a fraction of a day (a number), a deliberate, documented divergence
// from python-calamine (which yields a time object) — the shared model has no time type.
function parseOdsTime(value: string | undefined): number | undefined {
	if (value === undefined) return undefined;
	const m = /^-?P(?:(\d+)D)?T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(value);
	if (m === null) return undefined;
	const days = m[1] === undefined ? 0 : Number(m[1]);
	const hours = m[2] === undefined ? 0 : Number(m[2]);
	const mins = m[3] === undefined ? 0 : Number(m[3]);
	const secs = m[4] === undefined ? 0 : Number(m[4]);
	if (days === 0 && m[2] === undefined && m[3] === undefined && m[4] === undefined)
		return undefined;
	const fraction = days + (hours * 3600 + mins * 60 + secs) / 86_400;
	return Number.isFinite(fraction) ? fraction : undefined;
}

// Project an ODF cell onto the shared Cell model. Returns undefined for a cell with no representable
// content (it is then never emitted — the key to making empty repeats free).
function decodeOdsCell(
	valueType: string | undefined,
	isError: boolean,
	attrs: Readonly<Record<string, string>>,
	text: string,
): DecodedCell | undefined {
	if (isError) {
		// A formula error (LibreOffice's calcext:value-type="error"). Surface the displayed error text.
		const shown = text !== "" ? text : (attrs["office:value"] ?? "#ERR");
		return { type: "error", value: shown };
	}
	switch (valueType) {
		case "float":
		case "percentage":
		case "currency": {
			const n = Number(attrs["office:value"]);
			return Number.isFinite(n) ? { type: "number", value: n } : undefined;
		}
		case "boolean": {
			const b = attrs["office:boolean-value"];
			if (b === "true") return { type: "boolean", value: true };
			if (b === "false") return { type: "boolean", value: false };
			return undefined;
		}
		case "date": {
			const d = parseOdsDate(attrs["office:date-value"]);
			if (d !== undefined) return { type: "date", value: d };
			return text !== "" ? { type: "string", value: text } : undefined;
		}
		case "time": {
			const frac = parseOdsTime(attrs["office:time-value"]);
			return frac !== undefined ? { type: "number", value: frac } : undefined;
		}
		case "string": {
			const s = attrs["office:string-value"] ?? text;
			return { type: "string", value: s };
		}
		default:
			// No value-type: a plain-text or empty cell. Text ⇒ string; nothing ⇒ empty (dropped).
			return text !== "" ? { type: "string", value: text } : undefined;
	}
}

/**
 * Parse the whole content.xml (every sheet lives in this one part) into typed cell tables.
 * Tolerant by construction: unparseable cells degrade to empty, and no throw ever escapes here —
 * container-level failures are raised by the caller (`openOds`).
 */
export function parseOdsContent(xml: string): OdsSheet[] {
	const sheets: OdsSheet[] = [];

	// Document-position gating. Sheets are the top-level <table:table> inside <office:spreadsheet>;
	// a nested table (tableDepth > 1, e.g. inside a text cell) is ignored so it can't corrupt the
	// parent sheet's row/column cursor.
	let spreadsheetDepth = 0;
	let tableDepth = 0;

	// DOCUMENT-WIDE materialized-cell budget (not per-sheet): the repeat-bomb cap has to span every
	// sheet, or a tiny content.xml with N bomb sheets would materialize N × MAX_ODS_CELLS and OOM.
	// Once spent, later cells (and their merges/hyperlinks, which each ride a materialized cell) are
	// dropped — the reader-degrades way.
	let totalCells = 0;

	// Current sheet accumulation.
	let name = "";
	let state: SheetState = "visible";
	let cells = new Map<string, Cell>();
	let merges: string[] = [];
	let hyperlinks: Hyperlink[] = [];
	let minCol = Number.POSITIVE_INFINITY;
	let minRow = Number.POSITIVE_INFINITY;
	let maxCol = 0;
	let maxRow = 0;
	// rowCursor = number of rows already consumed in this sheet (the next row is rowCursor + 1).
	let rowCursor = 0;

	// Current row accumulation.
	let rowRepeat = 1;
	let colCursor = 0; // columns already consumed in this row (next cell is colCursor + 1).
	let rowCells: PendingCell[] = [];

	// Current cell accumulation.
	let inCell = false;
	let cellCovered = false;
	let cellRepeat = 1;
	let cellSpanCols = 1;
	let cellSpanRows = 1;
	let cellValueType: string | undefined;
	let cellIsError = false;
	let cellAttrs: Readonly<Record<string, string>> = {};
	let cellText = "";
	let cellHref: string | undefined;
	let sawParagraph = false;
	let paraDepth = 0; // > 0 while inside a <text:p> (character data counts only there).

	const beginSheet = (attrs: Readonly<Record<string, string>>): void => {
		name = attrs["table:name"] ?? "";
		const vis = attrs["table:visibility"];
		state = vis === "collapse" || vis === "filter" ? "hidden" : "visible";
		cells = new Map();
		merges = [];
		hyperlinks = [];
		minCol = Number.POSITIVE_INFINITY;
		minRow = Number.POSITIVE_INFINITY;
		maxCol = 0;
		maxRow = 0;
		rowCursor = 0;
	};

	const endSheet = (): void => {
		const dimension =
			cells.size === 0
				? undefined
				: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`;
		sheets.push({
			name,
			visible: state === "visible",
			state,
			cells,
			merges,
			hyperlinks,
			dimension,
		});
	};

	const beginRow = (attrs: Readonly<Record<string, string>>): void => {
		rowRepeat = intAttr(attrs["table:number-rows-repeated"], 1);
		colCursor = 0;
		rowCells = [];
	};

	const endRow = (): void => {
		if (rowCells.length > 0 && rowCursor < MAX_ROW && totalCells < MAX_ODS_CELLS) {
			const instances = Math.min(rowRepeat, MAX_ROW - rowCursor);
			for (let r = 0; r < instances && totalCells < MAX_ODS_CELLS; r++) {
				const row = rowCursor + 1 + r;
				for (const pc of rowCells) {
					if (pc.col > MAX_COL) continue;
					if (totalCells >= MAX_ODS_CELLS) break;
					const ref = formatRef({ col: pc.col, row });
					cells.set(ref, { ref, type: pc.decoded.type, value: pc.decoded.value } as Cell);
					totalCells += 1;
					if (pc.col < minCol) minCol = pc.col;
					if (pc.col > maxCol) maxCol = pc.col;
					if (row < minRow) minRow = row;
					if (row > maxRow) maxRow = row;
					// Merges + hyperlinks belong to the first row instance only (a repeated row carrying
					// a merge is degenerate; the first placement is the faithful one).
					if (r === 0 && pc.href !== undefined) {
						hyperlinks.push({ ref, target: pc.href });
					}
					if (r === 0 && (pc.spanCols > 1 || pc.spanRows > 1)) {
						const brCol = Math.min(pc.col + pc.spanCols - 1, MAX_COL);
						const brRow = Math.min(row + pc.spanRows - 1, MAX_ROW);
						// Skip a span that the grid-edge clamp collapsed to a single cell — a "XFD1:XFD1"
						// range isn't a merge, and the writer rightly rejects it.
						if (brCol > pc.col || brRow > row) {
							merges.push(`${ref}:${formatRef({ col: brCol, row: brRow })}`);
						}
					}
				}
			}
		}
		// Advance the row cursor by the full repeat (saturated) even when nothing was emitted, so a
		// trailing empty repeated row costs O(1) and later rows land at the right index.
		rowCursor = Math.min(rowCursor + rowRepeat, MAX_ROW);
	};

	const beginCell = (elementName: string, attrs: Readonly<Record<string, string>>): void => {
		cellCovered = elementName === "table:covered-table-cell";
		cellRepeat = intAttr(attrs["table:number-columns-repeated"], 1);
		cellSpanCols = intAttr(attrs["table:number-columns-spanned"], 1);
		cellSpanRows = intAttr(attrs["table:number-rows-spanned"], 1);
		cellValueType = attrs["office:value-type"];
		cellIsError = attrs["calcext:value-type"] === "error";
		cellAttrs = attrs;
		cellText = "";
		cellHref = undefined;
		sawParagraph = false;
		paraDepth = 0;
	};

	const endCell = (): void => {
		// A covered cell holds no value; both it and any empty cell only advance the column cursor.
		const decoded = cellCovered
			? undefined
			: decodeOdsCell(cellValueType, cellIsError, cellAttrs, cellText);
		if (decoded !== undefined) {
			const firstCol = colCursor + 1;
			const instances = Math.min(cellRepeat, Math.max(0, MAX_COL - colCursor));
			for (let k = 0; k < instances; k++) {
				rowCells.push({
					col: firstCol + k,
					decoded,
					// A hyperlink/merge applies to the first instance only.
					href: k === 0 ? cellHref : undefined,
					spanCols: k === 0 ? cellSpanCols : 1,
					spanRows: k === 0 ? cellSpanRows : 1,
				});
			}
		}
		colCursor = Math.min(colCursor + cellRepeat, MAX_COL);
		inCell = false;
	};

	const appendText = (s: string): void => {
		if (cellText.length >= MAX_CELL_TEXT) return;
		cellText += s;
		if (cellText.length > MAX_CELL_TEXT) cellText = cellText.slice(0, MAX_CELL_TEXT);
	};

	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const n = token.name;
			if (n === "office:spreadsheet") {
				if (!token.selfClosing) spreadsheetDepth++;
			} else if (n === "table:table") {
				if (spreadsheetDepth > 0) {
					tableDepth++;
					if (tableDepth === 1) {
						beginSheet(token.attrs);
						if (token.selfClosing) {
							endSheet();
							tableDepth--;
						}
					} else if (token.selfClosing) {
						tableDepth--;
					}
				}
			} else if (spreadsheetDepth > 0 && tableDepth === 1) {
				if (n === "table:table-row") {
					// A self-closed empty row (`<table:table-row/>`) carries no cells but still occupies
					// its repeat count of vertical space — advance the cursor so later rows land right.
					if (token.selfClosing) {
						rowCursor = Math.min(
							rowCursor + intAttr(token.attrs["table:number-rows-repeated"], 1),
							MAX_ROW,
						);
					} else {
						beginRow(token.attrs);
					}
				} else if (n === "table:table-cell" || n === "table:covered-table-cell") {
					beginCell(n, token.attrs);
					if (token.selfClosing) endCell();
					else inCell = true;
				} else if (inCell) {
					if (n === "text:p") {
						if (sawParagraph) appendText("\n");
						sawParagraph = true;
						if (!token.selfClosing) paraDepth++;
					} else if (paraDepth > 0) {
						if (n === "text:s") {
							const c = intAttr(token.attrs["text:c"], 1);
							appendText(" ".repeat(Math.min(c, MAX_CELL_TEXT)));
						} else if (n === "text:tab") {
							appendText("\t");
						} else if (n === "text:line-break") {
							appendText("\n");
						} else if (n === "text:a") {
							if (cellHref === undefined) cellHref = token.attrs["xlink:href"];
							// keep accumulating the link's inner text as the cell's text
						}
					}
				}
			}
		} else if (token.kind === "text") {
			if (inCell && paraDepth > 0) appendText(token.value);
		} else {
			// close
			const n = token.name;
			if (n === "office:spreadsheet") {
				if (spreadsheetDepth > 0) spreadsheetDepth--;
			} else if (n === "table:table") {
				if (tableDepth === 1) endSheet();
				if (tableDepth > 0) tableDepth--;
			} else if (spreadsheetDepth > 0 && tableDepth === 1) {
				if (n === "table:table-cell" || n === "table:covered-table-cell") {
					if (inCell) endCell();
				} else if (n === "table:table-row") {
					endRow();
				} else if (inCell && paraDepth > 0 && n === "text:p") {
					paraDepth--;
				}
			}
		}
	}

	return sheets;
}
