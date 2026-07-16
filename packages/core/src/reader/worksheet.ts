import { XlsxError } from "../errors";
import {
	type CellRef,
	type DecodeContext,
	decodeCell,
	decodeXstring,
	formatRef,
	MAX_COL,
	MAX_COL_WIDTH,
	MAX_FORMULA_LEN,
	MAX_PAGE_MARGIN,
	MAX_PAGE_SCALE,
	MAX_ROW,
	MAX_ROW_HEIGHT,
	MAX_SPIN_COUNT,
	MIN_PAGE_SCALE,
	parseCanonicalRange,
	parseRef,
	type Relationship,
	translateFormula,
} from "../ooxml";
import type {
	Cell,
	ColumnProps,
	Comment,
	FreezePane,
	HeaderFooter,
	Hyperlink,
	PageMargins,
	PageSetup,
	PrintOptions,
	Row,
	RowProps,
	SheetAutoFilter,
	SheetProtection,
} from "../types";
import { localName, relationshipId } from "../utils";
import { createXmlStream, tokenize, type XmlToken } from "../xml";

// `Row` is part of the public reader surface; its canonical home is ../types (shared across
// formats, M7). Re-exported here so the worksheet parsers and reader/workbook.ts keep their imports.
export type { Row } from "../types";

// Turn a worksheet part (xl/worksheets/sheetN.xml) into rows of typed cells. We walk the
// tokenizer event stream rather than building a DOM, so peak memory tracks one row, not the
// whole sheet. The same row state machine drives both the in-memory `readRows` (over a full
// string) and the constant-memory `streamRows` (over decompressed chunks) — see F2.2.
//
// Cells and rows are sparse and may appear out of order, so each cell carries its own A1
// ref and callers key by that ref — never by position in the array. When a `<c>` omits its
// `r` attribute (some streaming writers do), the column is taken positionally: one past the
// previous cell, exactly as the spec prescribes.
//
// Robustness mirrors the lower layers: the tokenizer never throws and does not validate
// structure, so this code must not corrupt a well-formed row because of a malformed
// neighbour. A new `<row>`/`<c>` opening while one is still unclosed finalizes the open one
// first, and a row left open at end-of-input (truncated file) is still emitted.

function safeColumn(ref: string): number | undefined {
	try {
		return parseRef(ref).col;
	} catch {
		return undefined;
	}
}

// A cell's format can be set at three levels, not just on the cell. `<c s>` is per-cell, but a
// whole column can be formatted with `<col … style>` and a whole row with `<row s customFormat>`.
// The effective style index resolves in this precedence (SpreadsheetML §18.3.1.13/.73/.4):
//
//     cell's own `s`  >  row default (customFormat)  >  column default (<col>)  >  style 0
//
// Without this, a date column whose cells omit `s` reads as plain numbers, and numberFormat()
// returns the wrong code — the common case where a producer formats a column once instead of
// per cell. All three levels are indices into the SAME cellXfs table, so the resolved index
// drops straight into StyleTable.isDateStyle/formatCode with no further translation.

// A column-default style: `<col … style>` applies to every cell in columns [min, max] that does
// not set its own `s`. (`style` is the column's default cell format — an index into cellXfs.)
interface ColumnStyle {
	readonly min: number;
	readonly max: number;
	readonly style: number;
}

function columnStyleFromToken(attrs: Readonly<Record<string, string>>): ColumnStyle | undefined {
	// A <col> with no `style` sets width/visibility only — it contributes no default format.
	if (attrs.style === undefined) return undefined;
	const style = Number(attrs.style);
	const min = Number(attrs.min);
	const max = Number(attrs.max);
	if (!Number.isInteger(style) || !Number.isInteger(min) || !Number.isInteger(max))
		return undefined;
	return { min, max, style };
}

// A row's default style is honored ONLY when `customFormat="1"`. The spec is explicit: `<row>`'s
// `s` is "only applied if the customFormat attribute is '1'". Writers routinely emit a bookkeeping
// `s` on ordinary rows, so without this gate we would restyle cells the file never meant to format.
function rowDefaultStyleFromToken(attrs: Readonly<Record<string, string>>): number | undefined {
	if (attrs.customFormat !== "1" && attrs.customFormat !== "true") return undefined;
	if (attrs.s === undefined) return undefined;
	const s = Number(attrs.s);
	return Number.isInteger(s) ? s : undefined;
}

// Resolve the precedence above for one cell. A present-but-unparseable cell `s` still counts as
// "the cell declared its own style" — it opts out of inherited defaults and falls back to style 0,
// rather than silently borrowing the row/column format.
function effectiveStyle(
	ownS: string | undefined,
	rowDefault: number | undefined,
	col: number | undefined,
	columns: readonly ColumnStyle[],
): number | undefined {
	if (ownS !== undefined) {
		const s = Number(ownS);
		return Number.isInteger(s) ? s : undefined;
	}
	if (rowDefault !== undefined) return rowDefault;
	if (col !== undefined) {
		// First covering range wins. Ranges don't overlap in a valid file; first-match keeps
		// resolution deterministic even if a malformed file declares a column twice.
		const range = columns.find((c) => col >= c.min && col <= c.max);
		if (range !== undefined) return range.style;
	}
	return undefined;
}

interface RowAssembler {
	/** Advance the state machine by one token; returns any rows that completed. */
	push(token: XmlToken): Row[];
	/** Emit a row left open at end of input (truncated file). */
	flush(): Row[];
}

function createRowAssembler(ctx: DecodeContext): RowAssembler {
	let inSheetData = false;
	let inRow = false;
	let lastRow = 0;
	let rowIndex = 0;
	let cells: Cell[] = [];
	let lastCol = 0;

	// Column defaults from `<cols>` (which precedes `<sheetData>`, so it is fully seen before any
	// cell) and the current row's default (customFormat) — both feed effectiveStyle for cells
	// that omit their own `s`. See the ColumnStyle/effectiveStyle notes above.
	const columns: ColumnStyle[] = [];
	let rowDefault: number | undefined;

	let inCell = false;
	let cellRef = "";
	let cellType: string | undefined;
	let cellStyle: number | undefined; // the `s` attribute (index into cellXfs)
	let cellIsInline = false; // type === 'inlineStr': value lives in <is>, not <v>
	let cellValue = "";
	let tBuf = ""; // raw text of the currently-open inline <t>, decoded as one unit on close
	let hasValue = false; // the value channel's element was present (even if empty)
	let inValue = false; // inside <v>
	let inInline = false; // inside <is>
	let textDepth = 0; // open <t> within <is>
	let phoneticDepth = 0; // open <rPh>/<phoneticPr> within <is> (excluded from the value)

	// An inline <t>'s raw text buffers until the element closes and decodes as ONE ST_Xstring
	// escape context (F9.6) — per run, like shared strings, so an escape never straddles runs.
	// Also flushed at cell finalize, so misnested markup (a missing </t>) drops no text.
	const flushT = () => {
		if (tBuf !== "") {
			cellValue += decodeXstring(tBuf);
			tBuf = "";
		}
	};

	// Finalize the open cell into the current row. A no-op when no cell is open.
	const flushCell = () => {
		if (!inCell) return;
		flushT();
		// A cached formula string (`t="str"`, its <v> is ST_Xstring) decodes as one context —
		// the single <v> element. Other <v> channels (numbers, booleans, errors, shared-string
		// INDEXES) are not xstrings and stay verbatim.
		const value = cellType === "str" ? decodeXstring(cellValue) : cellValue;
		cells.push(
			decodeCell(
				{
					ref: cellRef,
					type: cellType,
					value: hasValue ? value : undefined,
					style: cellStyle,
				},
				ctx,
			),
		);
		inCell = false;
	};

	function push(token: XmlToken): Row[] {
		const out: Row[] = [];

		if (token.kind === "open") {
			const name = localName(token.name);

			if (name === "sheetData") {
				if (!token.selfClosing) inSheetData = true;
				return out;
			}
			// `<col>` lives in `<cols>`, a sibling that precedes `<sheetData>` — capture its
			// default style before the inSheetData guard below would skip it.
			if (name === "col") {
				const cs = columnStyleFromToken(token.attrs);
				if (cs !== undefined) columns.push(cs);
				return out;
			}
			if (!inSheetData) return out;

			if (name === "row") {
				flushCell();
				if (inRow) out.push({ index: rowIndex, cells });
				const r = token.attrs.r;
				const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN;
				rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
				lastRow = rowIndex;
				// A non-customFormat row overwrites this back to undefined — no stale carryover.
				rowDefault = rowDefaultStyleFromToken(token.attrs);
				cells = [];
				lastCol = 0;
				if (token.selfClosing) {
					out.push({ index: rowIndex, cells });
					inRow = false;
				} else {
					inRow = true;
				}
				return out;
			}
			if (!inRow) return out;

			if (name === "c") {
				flushCell();
				const r = token.attrs.r;
				// The cell's column index, needed to resolve a `<col>` default style. Known even
				// when `r` is absent (positional: one past the previous cell).
				let col: number | undefined;
				if (r !== undefined) {
					cellRef = r;
					col = safeColumn(r);
					if (col !== undefined) lastCol = col;
				} else {
					lastCol += 1;
					col = lastCol;
					cellRef = formatRef({ col: lastCol, row: rowIndex });
				}
				cellType = token.attrs.t;
				// Effective style: cell `s` → row default → column default → style 0.
				cellStyle = effectiveStyle(token.attrs.s, rowDefault, col, columns);
				cellIsInline = cellType === "inlineStr";
				cellValue = "";
				tBuf = "";
				hasValue = false;
				inValue = false;
				inInline = false;
				textDepth = 0;
				phoneticDepth = 0;
				if (token.selfClosing) {
					cells.push(
						decodeCell(
							{ ref: cellRef, type: cellType, value: undefined, style: cellStyle },
							ctx,
						),
					);
				} else {
					inCell = true;
				}
				return out;
			}
			if (!inCell) return out;

			// A cell's value lives in exactly one channel, picked by its type: inline strings
			// in <is>/<t>, everything else in <v>. Gate on the type so a stray element from
			// the other channel can't pollute the value. Mark the value present as soon as its
			// element opens, so an explicit but empty <v></v> or <is><t></t></is> reads as ""
			// rather than collapsing to a blank cell.
			if (cellIsInline) {
				if (name === "is") {
					hasValue = true;
					if (!token.selfClosing) inInline = true;
				} else if (name === "t") {
					if (inInline && !token.selfClosing) textDepth++;
				} else if (name === "rPh" || name === "phoneticPr") {
					if (inInline && !token.selfClosing) phoneticDepth++;
				}
			} else if (name === "v") {
				hasValue = true;
				if (!token.selfClosing) inValue = true;
			}
			return out;
		}

		if (token.kind === "text") {
			const collect = cellIsInline
				? inInline && textDepth > 0 && phoneticDepth === 0
				: inValue;
			if (inCell && collect) {
				if (cellIsInline) tBuf += token.value;
				else cellValue += token.value;
				hasValue = true;
			}
			return out;
		}

		// close
		const name = localName(token.name);
		if (name === "sheetData") {
			flushCell();
			if (inRow) {
				out.push({ index: rowIndex, cells });
				inRow = false;
			}
			inSheetData = false;
			return out;
		}
		if (name === "row") {
			if (inRow) {
				flushCell();
				out.push({ index: rowIndex, cells });
				inRow = false;
			}
			return out;
		}
		if (!inCell) return out;
		if (name === "c") flushCell();
		else if (name === "v") inValue = false;
		else if (name === "is") inInline = false;
		else if (name === "t") {
			if (textDepth > 0) {
				textDepth--;
				if (textDepth === 0) flushT();
			}
		} else if (name === "rPh" || name === "phoneticPr") {
			if (phoneticDepth > 0) phoneticDepth--;
		}
		return out;
	}

	function flush(): Row[] {
		flushCell();
		if (inRow) {
			inRow = false;
			return [{ index: rowIndex, cells }];
		}
		return [];
	}

	return { push, flush };
}

/**
 * The comments a worksheet carries. Comments live in a separate part (xl/commentsN.xml) that
 * pairs an `<authors>` list with a `<commentList>`; each `<comment ref authorId>` holds rich
 * text in `<text>`. We concatenate the `<t>` runs (dropping formatting, matching how shared
 * strings read; each run decodes its `_xHHHH_` escapes as one ST_Xstring context — F9.6) and
 * resolve the authorId against the authors list. A comment with no `ref`, or an authorId that
 * resolves to nothing, still yields text — the author is just omitted.
 */
export function parseComments(xml: string): Comment[] {
	const authors: string[] = [];
	const comments: Comment[] = [];
	let inAuthors = false;
	let authorText: string | undefined; // non-undefined while inside an <author>
	let current: { ref: string; authorId: number } | undefined; // inside a <comment>
	let inText = false; // inside the current comment's <text>
	let tDepth = 0; // open <t> within <text>
	let text = "";
	let tBuf = ""; // raw text of the currently-open <t>, decoded as one unit on close

	const flushT = () => {
		if (tBuf !== "") {
			text += decodeXstring(tBuf);
			tBuf = "";
		}
	};

	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			if (name === "authors") {
				if (!token.selfClosing) inAuthors = true;
			} else if (name === "author" && inAuthors) {
				if (token.selfClosing) authors.push("");
				else authorText = "";
			} else if (name === "comment") {
				const ref = token.attrs.ref;
				// A missing/empty/non-integer authorId is unresolved (-1) so the author is omitted,
				// rather than falling through Number('')===0 to author 0 — don't fabricate an
				// attribution the file never made.
				const rawId = token.attrs.authorId;
				const authorId =
					rawId !== undefined && rawId !== "" && Number.isInteger(Number(rawId))
						? Number(rawId)
						: -1;
				current = ref !== undefined && ref !== "" ? { ref, authorId } : undefined;
				text = "";
				tBuf = "";
				inText = false;
				tDepth = 0;
			} else if (name === "text" && current !== undefined) {
				if (!token.selfClosing) inText = true;
			} else if (name === "t" && inText) {
				if (!token.selfClosing) tDepth++;
			}
		} else if (token.kind === "text") {
			if (authorText !== undefined) authorText += token.value;
			else if (inText && tDepth > 0) tBuf += token.value;
		} else {
			const name = localName(token.name);
			if (name === "authors") inAuthors = false;
			else if (name === "author" && authorText !== undefined) {
				authors.push(authorText);
				authorText = undefined;
			} else if (name === "t") {
				if (tDepth > 0) {
					tDepth--;
					if (tDepth === 0) flushT();
				}
			} else if (name === "text") inText = false;
			else if (name === "comment" && current !== undefined) {
				flushT(); // a missing </t> (misnested markup) must not drop collected text
				const author = authors[current.authorId];
				comments.push({
					ref: current.ref,
					...(author !== undefined && author !== "" ? { author } : {}),
					text,
				});
				current = undefined;
			}
		}
	}
	return comments;
}

/**
 * The worksheet's declared used range — the `<dimension ref>` (e.g. "A1:E10", or a single
 * cell). Optional in OOXML: undefined when the producer omits it, or when it is empty. The
 * element sits near the top of the part, so the scan returns as soon as it is found.
 */
export function parseDimension(xml: string): string | undefined {
	for (const token of tokenize(xml)) {
		if (token.kind === "open" && localName(token.name) === "dimension") {
			const ref = token.attrs.ref;
			if (ref !== undefined && ref !== "") return ref;
		}
	}
	return undefined;
}

/**
 * Map each cell to its EFFECTIVE style index, for resolving per-cell number formats. Effective
 * means the same precedence date detection uses — cell `s`, else the row default (`<row s
 * customFormat>`), else the column default (`<col … style>`), else style 0. A cell that resolves
 * to no style is omitted (the lookup falls back to style 0 anyway).
 *
 * Addressing MUST match the row assembler so numberFormat() and cell()/rows() agree on the same
 * ref: an explicit `r` is used verbatim; a cell without `r` is positioned one past the previous
 * cell (`formatRef`), exactly as the assembler synthesizes its ref — otherwise a no-`r` cell that
 * inherits a column/row style would read as a date via cell() but as "General" via numberFormat().
 */
export function parseCellStyles(xml: string): Map<string, number> {
	const styles = new Map<string, number>();
	const columns: ColumnStyle[] = [];
	let rowDefault: number | undefined;
	let lastRow = 0;
	let rowIndex = 0;
	let lastCol = 0;
	for (const token of tokenize(xml)) {
		if (token.kind !== "open") continue;
		const name = localName(token.name);
		if (name === "col") {
			const cs = columnStyleFromToken(token.attrs);
			if (cs !== undefined) columns.push(cs);
			continue;
		}
		if (name === "row") {
			const r = token.attrs.r;
			const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN;
			rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
			lastRow = rowIndex;
			rowDefault = rowDefaultStyleFromToken(token.attrs);
			lastCol = 0;
			continue;
		}
		if (name !== "c") continue;
		// Resolve the cell's ref and column with the same positional rule as the assembler.
		const r = token.attrs.r;
		let ref: string;
		let col: number | undefined;
		if (r !== undefined) {
			ref = r;
			col = safeColumn(r);
			if (col !== undefined) lastCol = col;
		} else {
			lastCol += 1;
			col = lastCol;
			ref = formatRef({ col: lastCol, row: rowIndex });
		}
		const index = effectiveStyle(token.attrs.s, rowDefault, col, columns);
		if (index !== undefined) styles.set(ref, index);
	}
	return styles;
}

/**
 * The formula text of every cell that carries one, keyed by A1 ref (F5.4) — a lazy dedicated scan
 * over `<f>` elements, addressing cells with the same positional rule as parseCellStyles. Plain and
 * array-master formulas are returned verbatim; a SHARED dependent (`<f t="shared" si/>` with no text
 * of its own) gets the master's text translated by the dependent's offset from the master.
 * `dataTable` formulas carry no reusable text and are skipped (a documented degradation).
 */
export function parseFormulas(xml: string): Map<string, string> {
	const formulas = new Map<string, string>();
	const masters = new Map<string, { anchor: CellRef; text: string }>(); // si → master
	const deps: { ref: string; cell: CellRef; si: string }[] = []; // shared dependents to resolve
	let inSheetData = false;
	let rowIndex = 0;
	let lastRow = 0;
	let lastCol = 0;
	let curRef: string | undefined;
	let curCell: CellRef | undefined;
	let inF = false;
	let fType: string | undefined;
	let fSi: string | undefined;
	let fText = "";

	// Record a completed <f> for the current cell. Empty text ⟺ a shared dependent (its text lives
	// on the master); non-empty text is a plain, array-master, or shared-master formula.
	const finalize = (text: string): void => {
		if (curRef === undefined) return;
		if (fType === "shared") {
			if (fSi === undefined) return;
			if (text === "") {
				if (curCell !== undefined) deps.push({ ref: curRef, cell: curCell, si: fSi });
			} else {
				// Register the master for translation ONLY when its text is within Excel's length
				// ceiling. A hostile file with an over-long master would otherwise make the second pass
				// O(dependents × masterLength) — file-size-quadratic; capping keeps each translation
				// O(MAX_FORMULA_LEN). Over-long masters degrade (their dependents get no formula).
				if (curCell !== undefined && text.length <= MAX_FORMULA_LEN) {
					masters.set(fSi, { anchor: curCell, text });
				}
				formulas.set(curRef, text); // the master's own formula is its text
			}
		} else if (fType === "dataTable") {
			// No reusable formula text — degrade to the cached value only.
		} else if (text !== "") {
			formulas.set(curRef, text); // plain or array-master formula, verbatim
		}
	};

	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			// Only formulas INSIDE <sheetData> are cells. A stray <c><f> in an oleObjects /
			// AlternateContent block (or anywhere else) must not fabricate a formula on a real cell —
			// this mirrors the row assembler's own sheetData gate.
			if (name === "sheetData") {
				if (!token.selfClosing) inSheetData = true;
				continue;
			}
			if (!inSheetData) continue;
			if (name === "row") {
				const r = token.attrs.r;
				const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN;
				rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
				lastRow = rowIndex;
				lastCol = 0;
			} else if (name === "c") {
				const r = token.attrs.r;
				if (r !== undefined) {
					curRef = r;
					const col = safeColumn(r);
					curCell = col !== undefined ? { col, row: rowIndex } : undefined;
					if (col !== undefined) lastCol = col;
				} else {
					lastCol += 1;
					curCell = { col: lastCol, row: rowIndex };
					curRef = formatRef(curCell);
				}
			} else if (name === "f") {
				fType = token.attrs.t;
				fSi = token.attrs.si;
				fText = "";
				if (token.selfClosing) finalize("");
				else inF = true;
			}
		} else if (token.kind === "text") {
			if (inF) fText += token.value;
		} else {
			const name = localName(token.name);
			if (name === "sheetData") inSheetData = false;
			else if (inF && name === "f") {
				finalize(fText);
				inF = false;
			}
		}
	}

	// Second pass: a dependent's formula is the master's text shifted by the dependent's offset.
	for (const d of deps) {
		const master = masters.get(d.si);
		if (master === undefined) continue;
		formulas.set(
			d.ref,
			translateFormula(
				master.text,
				d.cell.row - master.anchor.row,
				d.cell.col - master.anchor.col,
			),
		);
	}
	return formulas;
}

/** Read rows from a fully in-memory worksheet string. */
export function* readRows(xml: string, ctx: DecodeContext): Generator<Row> {
	const assembler = createRowAssembler(ctx);
	for (const token of tokenize(xml)) yield* assembler.push(token);
	yield* assembler.flush();
}

/**
 * The merged-cell ranges a worksheet declares, in document order, as A1 ranges (e.g. `A1:B2`).
 * They live in a `<mergeCells>` block after `<sheetData>`, so we scan the token stream for the
 * `<mergeCell ref>` children rather than the row state machine. Refs are returned verbatim as
 * the producer wrote them; a missing or empty `ref` is skipped.
 */
export function parseMergedCells(xml: string): string[] {
	const ranges: string[] = [];
	for (const token of tokenize(xml)) {
		if (token.kind === "open" && localName(token.name) === "mergeCell") {
			const ref = token.attrs.ref;
			if (ref !== undefined && ref !== "") ranges.push(ref);
		}
	}
	return ranges;
}

/**
 * The sheet's autoFilter range (filter dropdowns), or `undefined` when the sheet declares none (F10.2).
 * ONLY an `<autoFilter>` that is a DIRECT child of `<worksheet>` is the sheet-level filter: `<autoFilter>`
 * also nests inside `<customSheetViews>/<customSheetView>` (a saved view that retained a filter), and a
 * flat scan would surface that as an active filter the sheet doesn't have — the exact trap parseFreezePane
 * was hardened against for `<pane>`. So the scan tracks nesting depth and accepts the filter only at
 * depth 1 (worksheet the sole open ancestor). The range is validated as a canonical, in-grid A1 range and
 * kept SYMBOLIC (never expanded per-cell — F4.4/F4.6); a hostile or non-canonical ref is DROPPED (the
 * strict writer would reject it, and real producers write canonical), so what the reader returns is
 * always writable.
 */
export function parseAutoFilter(xml: string): SheetAutoFilter | undefined {
	let depth = 0; // number of currently-open ancestor elements
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			// A direct child of <worksheet> has worksheet as its only open ancestor → depth === 1. A
			// customSheetView's nested <autoFilter> sits at depth 3 and is skipped.
			if (depth === 1 && localName(token.name) === "autoFilter") {
				const ref = token.attrs.ref;
				return ref !== undefined && parseCanonicalRange(ref) !== undefined
					? { ref }
					: undefined;
			}
			if (!token.selfClosing) depth++;
		} else if (token.kind === "close") {
			depth--;
		}
	}
	return undefined;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

function bool01(v: string | undefined): boolean | undefined {
	if (v === "1" || v === "true") return true;
	if (v === "0" || v === "false") return false;
	return undefined;
}

// The boolean attributes of <sheetProtection>, all optional and carried verbatim (F10.3).
const SHEET_PROTECTION_FLAGS = [
	"sheet",
	"objects",
	"scenarios",
	"formatCells",
	"formatColumns",
	"formatRows",
	"insertColumns",
	"insertRows",
	"insertHyperlinks",
	"deleteColumns",
	"deleteRows",
	"selectLockedCells",
	"selectUnlockedCells",
	"sort",
	"autoFilter",
	"pivotTables",
] as const;

/**
 * The sheet's `<sheetProtection>` (F10.3), or `undefined` when it declares none. Scoped to a DIRECT
 * child of `<worksheet>` (depth 1) — the same guard as {@link parseAutoFilter}, since `<sheetProtection>`
 * is worksheet-level and must not be confused with any look-alike deeper in the part. Every present
 * boolean attribute and any password material (`password`, `algorithmName`/`hashValue`/`saltValue`/
 * `spinCount`) is kept VERBATIM — openjsxl never computes or verifies a hash; the writer re-emits it.
 */
export function parseSheetProtection(xml: string): SheetProtection | undefined {
	let depth = 0;
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			if (depth === 1 && localName(token.name) === "sheetProtection") {
				return readSheetProtection(token.attrs);
			}
			if (!token.selfClosing) depth++;
		} else if (token.kind === "close") {
			depth--;
		}
	}
	return undefined;
}

function readSheetProtection(attrs: Readonly<Record<string, string>>): SheetProtection | undefined {
	const out: Mutable<SheetProtection> = {};
	for (const flag of SHEET_PROTECTION_FLAGS) {
		const b = bool01(attrs[flag]);
		if (b !== undefined) out[flag] = b;
	}
	if (attrs.password !== undefined) out.password = attrs.password;
	if (attrs.algorithmName !== undefined) out.algorithmName = attrs.algorithmName;
	if (attrs.hashValue !== undefined) out.hashValue = attrs.hashValue;
	if (attrs.saltValue !== undefined) out.saltValue = attrs.saltValue;
	// spinCount is xsd:unsignedInt — DROP an out-of-range value (a 21-digit count would otherwise become
	// a float that re-emits as `1e+21`, invalid integer XML). Shared bound with the writer's reject.
	const spin = attrs.spinCount;
	if (spin !== undefined && /^[0-9]+$/.test(spin)) {
		const n = Number(spin);
		if (Number.isInteger(n) && n <= MAX_SPIN_COUNT) out.spinCount = n;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

// ── Print setup (F10.4): printOptions, pageMargins, pageSetup, headerFooter ────────────────────────

/** One pass over a worksheet's four print-setup elements (each absent when the sheet declares none). */
export interface PrintSetup {
	readonly printOptions?: PrintOptions;
	readonly pageMargins?: PageMargins;
	readonly pageSetup?: PageSetup;
	readonly headerFooter?: HeaderFooter;
}

// A finite double attribute clamped into [0, max]; undefined when absent or non-numeric/non-finite.
function clampedDouble(v: string | undefined, max: number): number | undefined {
	if (v === undefined) return undefined;
	const n = Number(v);
	if (!Number.isFinite(n)) return undefined;
	return n < 0 ? 0 : n > max ? max : n;
}

// A non-negative integer attribute within the xsd:unsignedInt ceiling; undefined otherwise. Gated to
// canonical digits first so a hostile 21-digit value is DROPPED, not coerced to a lossy float.
function uintAttr(v: string | undefined): number | undefined {
	if (v === undefined || !/^[0-9]+$/.test(v)) return undefined;
	const n = Number(v);
	return Number.isInteger(n) && n <= 0xffffffff ? n : undefined;
}

// An enum attribute restricted to a fixed set; undefined when absent/unrecognized (falls to spec default).
function enumAttr<T extends string>(v: string | undefined, allowed: readonly T[]): T | undefined {
	return v !== undefined && (allowed as readonly string[]).includes(v) ? (v as T) : undefined;
}

function readPrintOptions(attrs: Readonly<Record<string, string>>): PrintOptions | undefined {
	const out: Mutable<PrintOptions> = {};
	for (const flag of [
		"gridLines",
		"headings",
		"horizontalCentered",
		"verticalCentered",
	] as const) {
		const b = bool01(attrs[flag]);
		if (b !== undefined) out[flag] = b;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

function readPageMargins(attrs: Readonly<Record<string, string>>): PageMargins | undefined {
	// All six are required by the schema — produce a value only when every one is a finite number
	// (clamped to [0, MAX_PAGE_MARGIN]); a missing/non-finite margin drops the whole element.
	const left = clampedDouble(attrs.left, MAX_PAGE_MARGIN);
	const right = clampedDouble(attrs.right, MAX_PAGE_MARGIN);
	const top = clampedDouble(attrs.top, MAX_PAGE_MARGIN);
	const bottom = clampedDouble(attrs.bottom, MAX_PAGE_MARGIN);
	const header = clampedDouble(attrs.header, MAX_PAGE_MARGIN);
	const footer = clampedDouble(attrs.footer, MAX_PAGE_MARGIN);
	if (
		left === undefined ||
		right === undefined ||
		top === undefined ||
		bottom === undefined ||
		header === undefined ||
		footer === undefined
	) {
		return undefined;
	}
	return { left, right, top, bottom, header, footer };
}

function readPageSetup(attrs: Readonly<Record<string, string>>): PageSetup | undefined {
	const out: Mutable<PageSetup> = {};
	const paperSize = uintAttr(attrs.paperSize);
	if (paperSize !== undefined) out.paperSize = paperSize;
	const orientation = enumAttr(attrs.orientation, ["default", "portrait", "landscape"] as const);
	if (orientation !== undefined) out.orientation = orientation;
	// scale is a percentage in [MIN_PAGE_SCALE, MAX_PAGE_SCALE] — clamp a finite integer into range.
	if (attrs.scale !== undefined && /^[0-9]+$/.test(attrs.scale)) {
		const s = Number(attrs.scale);
		if (Number.isInteger(s)) {
			out.scale =
				s < MIN_PAGE_SCALE ? MIN_PAGE_SCALE : s > MAX_PAGE_SCALE ? MAX_PAGE_SCALE : s;
		}
	}
	const fitToWidth = uintAttr(attrs.fitToWidth);
	if (fitToWidth !== undefined) out.fitToWidth = fitToWidth;
	const fitToHeight = uintAttr(attrs.fitToHeight);
	if (fitToHeight !== undefined) out.fitToHeight = fitToHeight;
	const firstPageNumber = uintAttr(attrs.firstPageNumber);
	if (firstPageNumber !== undefined) out.firstPageNumber = firstPageNumber;
	const useFirstPageNumber = bool01(attrs.useFirstPageNumber);
	if (useFirstPageNumber !== undefined) out.useFirstPageNumber = useFirstPageNumber;
	const blackAndWhite = bool01(attrs.blackAndWhite);
	if (blackAndWhite !== undefined) out.blackAndWhite = blackAndWhite;
	const draft = bool01(attrs.draft);
	if (draft !== undefined) out.draft = draft;
	const cellComments = enumAttr(attrs.cellComments, ["none", "asDisplayed", "atEnd"] as const);
	if (cellComments !== undefined) out.cellComments = cellComments;
	const pageOrder = enumAttr(attrs.pageOrder, ["downThenOver", "overThenDown"] as const);
	if (pageOrder !== undefined) out.pageOrder = pageOrder;
	return Object.keys(out).length > 0 ? out : undefined;
}

const HF_CHILDREN = [
	"oddHeader",
	"oddFooter",
	"evenHeader",
	"evenFooter",
	"firstHeader",
	"firstFooter",
] as const;

function readHeaderFooterAttrs(attrs: Readonly<Record<string, string>>): Mutable<HeaderFooter> {
	const out: Mutable<HeaderFooter> = {};
	for (const flag of [
		"differentOddEven",
		"differentFirst",
		"scaleWithDoc",
		"alignWithMargins",
	] as const) {
		const b = bool01(attrs[flag]);
		if (b !== undefined) out[flag] = b;
	}
	return out;
}

/**
 * Parse a worksheet's four print-setup elements (F10.4) in ONE token pass. Each of `printOptions`,
 * `pageMargins`, `pageSetup`, `headerFooter` is accepted only as a DIRECT child of `<worksheet>`
 * (depth 1) — a `<customSheetView>` carries its own copies of all four, and a flat scan would surface
 * those instead (the parseAutoFilter/parseFreezePane scoping precedent). The header/footer child strings
 * carry Excel's `&`-codes verbatim, decoded as one ST_Xstring context (like an inline `<t>`).
 */
export function parsePrintSetup(xml: string): PrintSetup {
	const out: Mutable<PrintSetup> = {};
	let depth = 0;
	let hf: Mutable<HeaderFooter> | undefined; // the depth-1 headerFooter being built
	let inHeaderFooter = false;
	let hfChild: (typeof HF_CHILDREN)[number] | undefined; // which child string is accumulating
	let hfText = "";
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			if (depth === 1) {
				if (name === "printOptions") {
					const po = readPrintOptions(token.attrs);
					if (po !== undefined) out.printOptions = po;
				} else if (name === "pageMargins") {
					const m = readPageMargins(token.attrs);
					if (m !== undefined) out.pageMargins = m;
				} else if (name === "pageSetup") {
					const ps = readPageSetup(token.attrs);
					if (ps !== undefined) out.pageSetup = ps;
				} else if (name === "headerFooter") {
					hf = readHeaderFooterAttrs(token.attrs);
					if (token.selfClosing) {
						if (Object.keys(hf).length > 0) out.headerFooter = hf;
						hf = undefined;
					} else {
						inHeaderFooter = true;
					}
				}
			} else if (
				inHeaderFooter &&
				depth === 2 &&
				(HF_CHILDREN as readonly string[]).includes(name)
			) {
				hfChild = name as (typeof HF_CHILDREN)[number];
				hfText = "";
			}
			if (!token.selfClosing) depth++;
		} else if (token.kind === "text") {
			if (hfChild !== undefined) hfText += token.value;
		} else {
			depth--;
			const cname = localName(token.name);
			if (hfChild !== undefined && cname === hfChild) {
				if (hfText.length > 0 && hf !== undefined) hf[hfChild] = decodeXstring(hfText);
				hfChild = undefined;
			} else if (inHeaderFooter && cname === "headerFooter") {
				if (hf !== undefined && Object.keys(hf).length > 0) out.headerFooter = hf;
				hf = undefined;
				inHeaderFooter = false;
			}
		}
	}
	return out;
}

/**
 * The hyperlinks a worksheet declares, in document order. They live in a `<hyperlinks>` block
 * (a sibling of `<sheetData>`); each `<hyperlink>` carries the covered `ref` plus some of:
 * an `r:id` pointing into the worksheet's own relationships (the external target), an
 * in-workbook `location`, a `tooltip`, and a `display` override. The external target is joined
 * from `rels` here so callers get the resolved URL, not a bare relationship id; a link with no
 * resolvable `r:id` simply has no `target`. A missing/empty `ref` is skipped.
 */
export function parseHyperlinks(xml: string, rels?: Map<string, Relationship>): Hyperlink[] {
	const links: Hyperlink[] = [];
	for (const token of tokenize(xml)) {
		if (token.kind !== "open" || localName(token.name) !== "hyperlink") continue;
		const ref = token.attrs.ref;
		if (ref === undefined || ref === "") continue;

		// Built immutably (Hyperlink is readonly): include only the attributes that are present.
		const rid = relationshipId(token.attrs);
		const target = rid !== undefined ? rels?.get(rid)?.target : undefined;
		const location = token.attrs.location;
		const tooltip = token.attrs.tooltip;
		const display = token.attrs.display;
		links.push({
			ref,
			// An empty external target is no destination — gate it exactly like an empty location,
			// so a degenerate `Target=""` rel doesn't surface as a target the writer then melts
			// away (which would make read→write→read lossy instead of lossless-or-typed).
			...(target !== undefined && target !== "" ? { target } : {}),
			...(location !== undefined && location !== "" ? { location } : {}),
			...(tooltip !== undefined ? { tooltip } : {}),
			...(display !== undefined ? { display } : {}),
		});
	}
	return links;
}

/**
 * Read rows from a stream of decompressed worksheet chunks without materializing the part —
 * peak memory tracks one row, not the whole sheet (F2.2). Bytes are decoded with a streaming
 * `TextDecoder` (multi-byte sequences may split across chunks) and fed through the chunk-safe
 * tokenizer.
 */
export async function* streamRows(
	chunks: AsyncIterable<Uint8Array>,
	ctx: DecodeContext,
): AsyncGenerator<Row> {
	const assembler = createRowAssembler(ctx);
	const xml = createXmlStream();
	const decoder = new TextDecoder();

	// A single cell VALUE that grows past the JS engine's maximum string length (V8 ~512 MiB) makes
	// the accumulation throw a bare `RangeError` — surface it as the typed `part-too-large` the
	// tolerant-reader contract requires, never a bare throw (F9.7 review follow-up). The typed
	// MAX_UNFINISHED_CONSTRUCT cap in the stream already passes through untouched.
	try {
		for await (const bytes of chunks) {
			const text = decoder.decode(bytes, { stream: true });
			if (text === "") continue;
			for (const token of xml.push(text)) yield* assembler.push(token);
		}
		const tail = decoder.decode(); // finalize any pending multi-byte sequence
		if (tail !== "") for (const token of xml.push(tail)) yield* assembler.push(token);
		for (const token of xml.flush()) yield* assembler.push(token);
		yield* assembler.flush();
	} catch (cause) {
		if (cause instanceof RangeError) {
			throw new XlsxError(
				"part-too-large",
				"a cell value exceeds the maximum string length this runtime supports",
				{ cause },
			);
		}
		throw cause;
	}
}

// ── Sheet geometry (F4.5) ──────────────────────────────────────────────────────────────────────
// Dedicated scans in the mergedCells idiom — lazy, off the hot row path. Like the style model,
// these DEGRADE anything outside the shared reader/writer bounds (grid limits, width/height
// ceilings), so whatever they return, the writer accepts — the bridge can carry geometry without
// ever crashing on a tolerated file.

const boolFlag = (val: string | undefined): boolean => val === "1" || val === "true";

/**
 * Column width/visibility declarations from `<cols>`, in document order. Entries carrying only a
 * default STYLE (no width, not hidden) are style plumbing, not geometry — they are omitted here
 * (style resolution already honors them). Ranges outside Excel's grid and out-of-ceiling widths
 * degrade to absent.
 */
export function parseColumnProps(xml: string): ColumnProps[] {
	const out: ColumnProps[] = [];
	for (const token of tokenize(xml)) {
		if (token.kind !== "open") continue;
		const name = localName(token.name);
		// <cols> always precedes <sheetData>; a col-named element after it (extension lists,
		// alternate content) is not column geometry — stop scanning at the cell region.
		if (name === "sheetData") break;
		if (name !== "col") continue;
		const min = Number(token.attrs.min);
		const max = Number(token.attrs.max);
		if (!Number.isInteger(min) || !Number.isInteger(max)) continue;
		if (min < 1 || max < min || max > MAX_COL) continue;
		const props: { min: number; max: number; width?: number; hidden?: boolean } = { min, max };
		if (token.attrs.width !== undefined) {
			const width = Number(token.attrs.width);
			if (Number.isFinite(width) && width > 0 && width <= MAX_COL_WIDTH) props.width = width;
		}
		if (boolFlag(token.attrs.hidden)) props.hidden = true;
		if (props.width !== undefined || props.hidden) out.push(props);
	}
	return out;
}

/**
 * Per-row height/visibility from `<row ht hidden>` attributes, keyed by 1-based row index (the
 * `r` attribute, or positional — one past the previous row — when absent, matching the row
 * assembler). Rows with neither property are absent; heights outside (0, 409.5] degrade.
 */
export function parseRowProperties(xml: string): Map<number, RowProps> {
	const out = new Map<number, RowProps>();
	let lastRow = 0;
	let inSheetData = false;
	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		const name = localName(token.name);
		if (token.kind === "close") {
			if (name === "sheetData") inSheetData = false;
			continue;
		}
		if (name === "sheetData") {
			if (!token.selfClosing) inSheetData = true;
			continue;
		}
		if (!inSheetData || name !== "row") continue;
		// Resolve the row index with the EXACT rule the row assembler uses (parseInt, positional
		// fallback) — Number() disagrees with parseInt on tolerated-malformed values ("1e3",
		// "3abc"), which would attach a height to a different row than its cells and cascade the
		// positional counter for every r-less row after it (adversarial review, F4.5).
		const rAttr = token.attrs.r;
		const parsed = rAttr !== undefined ? Number.parseInt(rAttr, 10) : Number.NaN;
		const rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
		lastRow = rowIndex;
		if (rowIndex > MAX_ROW) continue;
		const props: { height?: number; hidden?: boolean } = {};
		if (token.attrs.ht !== undefined) {
			const height = Number(token.attrs.ht);
			if (Number.isFinite(height) && height > 0 && height <= MAX_ROW_HEIGHT) {
				props.height = height;
			}
		}
		if (boolFlag(token.attrs.hidden)) props.hidden = true;
		if (props.height !== undefined || props.hidden) out.set(rowIndex, props);
	}
	return out;
}

/**
 * The sheet's frozen pane from the `<sheetViews>` block's `<pane>`, or `undefined` when none.
 * Only `state="frozen"` is modelled: for frozen panes xSplit/ySplit are whole column/row counts,
 * while split (and frozenSplit) panes measure in twentieths of a point — a different world,
 * deferred and read as no freeze.
 *
 * The scan is SCOPED to `<sheetViews>`: `<pane>` also appears inside `<customSheetViews>` (a
 * saved Custom View, after `<sheetData>`), and picking that one up would fabricate a freeze the
 * active view doesn't have (adversarial review, F4.5).
 */
export function parseFreezePane(xml: string): FreezePane | undefined {
	let inSheetViews = false;
	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		const name = localName(token.name);
		if (name === "sheetViews") {
			// The block closed (or was empty) without a pane: nothing is frozen. sheetViews
			// precedes sheetData, so stopping here also skips the whole cell region.
			if (token.kind === "close" || token.selfClosing) return undefined;
			inSheetViews = true;
			continue;
		}
		if (!inSheetViews || token.kind !== "open" || name !== "pane") continue;
		if (token.attrs.state !== "frozen") return undefined;
		const x = token.attrs.xSplit === undefined ? 0 : Number(token.attrs.xSplit);
		const y = token.attrs.ySplit === undefined ? 0 : Number(token.attrs.ySplit);
		const cols = Number.isInteger(x) && x > 0 && x < MAX_COL ? x : 0;
		const rows = Number.isInteger(y) && y > 0 && y < MAX_ROW ? y : 0;
		if (cols === 0 && rows === 0) return undefined;
		const out: { rows?: number; cols?: number } = {};
		if (rows > 0) out.rows = rows;
		if (cols > 0) out.cols = cols;
		return out;
	}
	return undefined;
}
