import { XlsxError } from "../errors";
import {
	type CellRef,
	type DecodeContext,
	decodeXstring,
	formatRef,
	MAX_COL,
	MAX_COL_WIDTH,
	MAX_FORMULA_LEN,
	MAX_ROW,
	MAX_ROW_HEIGHT,
	parseCanonicalRange,
	type Relationship,
	rangeRunsForward,
	translateFormula,
} from "../ooxml";
import type {
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
import { isXmlSafe, localName, relationshipId } from "../utils";
import { createXmlStream, tokenize } from "../xml";
import { createRowAssembler } from "./row-assembler";
import {
	type HeaderFooterChild,
	headerFooterFromAttributes,
	isHeaderFooterChild,
	type Mutable,
	pageMarginsFromAttributes,
	pageSetupFromAttributes,
	printOptionsFromAttributes,
	sheetProtectionFromAttributes,
	trueAttribute,
} from "./worksheet-attributes";
import { effectiveStyle, rowDefaultStyleFromToken, safeColumn } from "./worksheet-cell";
import { columnStyleFromToken, createColumnStyleResolver } from "./worksheet-column-styles";
import {
	createWorksheetStructure,
	isWorksheetRow,
	pushWorksheetClose,
	pushWorksheetOpen,
} from "./worksheet-structure";

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
				// Decode the complete ST_Xstring once, after all XML text tokens have been joined.
				authors.push(decodeXstring(authorText));
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
 * Only worksheet/cols and worksheet/sheetData/row contribute metadata; extension lookalikes are
 * outside that shared structural scope and cannot overwrite a real cell.
 */
export function parseCellStyles(xml: string): Map<string, number> {
	const styles = new Map<string, number>();
	const columns = createColumnStyleResolver();
	const structure = createWorksheetStructure();
	let rowDefault: number | undefined;
	let lastRow = 0;
	let rowIndex = 0;
	let lastCol = 0;
	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		if (token.kind === "close") {
			const ended = pushWorksheetClose(structure, token.name);
			if (ended === "sheetData" || ended === "row") {
				rowDefault = undefined;
				lastCol = 0;
			}
			continue;
		}

		const opened = pushWorksheetOpen(structure, token.name, token.selfClosing);
		if (opened === "col") {
			const cs = columnStyleFromToken(token.attrs);
			if (cs !== undefined) columns.add(cs);
			continue;
		}
		if (opened === "row") {
			const r = token.attrs.r;
			const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN;
			rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
			lastRow = rowIndex;
			rowDefault = isWorksheetRow(structure)
				? rowDefaultStyleFromToken(token.attrs)
				: undefined;
			lastCol = 0;
			continue;
		}
		if (opened !== "c") continue;
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
 *
 * The scan follows the shared structure tracker (F11.38): rows and cells advance only where the row
 * assembler admits them, and a formula is only an admitted cell's own direct `<f>`, so an `<f>` in an
 * extension, in a fake row or cell, or nested in other cell content supplies nothing, and text inside
 * an extension within the `<f>` is not part of it. An `<f>` its cell closes before `</f>` is dropped.
 *
 * @example parseFormulas('<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>') // Map { "A1" => "1+1" }
 * @example parseFormulas('<worksheet><sheetData><row r="1"><c r="A1"><v>2</v><extLst><ext><x:f>HACK()</x:f></ext></extLst></c></row></sheetData></worksheet>') // Map {} (the extension's <f> is not the cell's)
 */
export function parseFormulas(xml: string): Map<string, string> {
	const formulas = new Map<string, string>();
	const masters = new Map<string, { anchor: CellRef; text: string }>(); // si → master
	const deps: { ref: string; cell: CellRef; si: string }[] = []; // shared dependents to resolve
	const structure = createWorksheetStructure();
	let rowIndex = 0;
	let lastRow = 0;
	let lastCol = 0;
	let curRef: string | undefined;
	let curCell: CellRef | undefined;
	let fType: string | undefined;
	let fSi: string | undefined;
	let fText = "";

	// Record a completed <f> for the current cell. Empty text ⟺ a shared dependent (its text lives
	// on the master); non-empty text is a plain, array-master, or shared-master formula.
	const finalize = (raw: string): void => {
		if (curRef === undefined) return;
		// ST_Formula is a restriction of ST_Xstring (F11.20): decode the COMPLETE accumulated element
		// text exactly once, BEFORE the length ceiling, master registration, and dependent translation,
		// so the accessor, the bridge, and the evaluator all see logical text (a stored `_x005F_x0041_`
		// IS the literal `_x0041_`; a raw `_x0041_` is the character `A`, as Excel reads it). A decoded
		// value the strict writer cannot emit (a control char / lone surrogate a producer escaped)
		// degrades to "no formula" — the cached value survives — keeping the shared bound closed.
		const logical = raw === "" ? "" : decodeXstring(raw);
		// Stored form carries no leading `=` (the writer refuses one): a producer that wrote `=A1` —
		// raw, or as the escape `_x003D_A1` — meant the formula A1, so strip one, the same
		// normalization the DV/CF parsers apply. A dependent is recognized by its EMPTY wire text.
		const text = logical.startsWith("=") ? logical.slice(1) : logical;
		if (raw !== "" && (text === "" || !isXmlSafe(text))) return;
		if (fType === "shared") {
			if (fSi === undefined) return;
			if (raw === "") {
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
		if (token.kind === "text") {
			// Only the cell's own open <f> collects text, and not inside an extension within it.
			if (structure.cellChild?.kind === "f" && structure.cellBarrier === undefined) {
				fText += token.value;
			}
			continue;
		}
		if (token.kind === "close") {
			if (pushWorksheetClose(structure, token.name) === "f") finalize(fText);
			continue;
		}
		// Only formulas of admitted cells INSIDE canonical <sheetData> are cells' formulas. A stray
		// <c><f> in an oleObjects / AlternateContent block, an extension, or a fake row must not
		// fabricate a formula on a real cell; this is the row assembler's own structural gate.
		const opened = pushWorksheetOpen(structure, token.name, token.selfClosing);
		if (opened === "row") {
			const r = token.attrs.r;
			const parsed = r !== undefined ? Number.parseInt(r, 10) : Number.NaN;
			rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : lastRow + 1;
			lastRow = rowIndex;
			lastCol = 0;
		} else if (opened === "c") {
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
		} else if (opened === "f") {
			fType = token.attrs.t;
			fSi = token.attrs.si;
			fText = "";
			if (token.selfClosing) finalize("");
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
				if (ref === undefined) return undefined;
				const range = parseCanonicalRange(ref);
				// Drop a non-canonical, out-of-grid, OR backwards (bottom-right before top-left) ref: the
				// strict writer rejects a backwards range, so dropping it keeps "what the reader returns
				// is always writable" true — the shared bound with the writer (rangeRunsForward).
				return range !== undefined && rangeRunsForward(range) ? { ref } : undefined;
			}
			if (!token.selfClosing) depth++;
		} else if (token.kind === "close") {
			depth--;
		}
	}
	return undefined;
}

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
				return sheetProtectionFromAttributes(token.attrs);
			}
			if (!token.selfClosing) depth++;
		} else if (token.kind === "close") {
			depth--;
		}
	}
	return undefined;
}

// ── Print setup (F10.4): printOptions, pageMargins, pageSetup, headerFooter ────────────────────────

/** One pass over a worksheet's four print-setup elements (each absent when the sheet declares none). */
export interface PrintSetup {
	readonly printOptions?: PrintOptions;
	readonly pageMargins?: PageMargins;
	readonly pageSetup?: PageSetup;
	readonly headerFooter?: HeaderFooter;
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
	let hfChild: HeaderFooterChild | undefined; // which child string is accumulating
	let hfText = "";
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			if (depth === 1) {
				if (name === "printOptions") {
					const po = printOptionsFromAttributes(token.attrs);
					if (po !== undefined) out.printOptions = po;
				} else if (name === "pageMargins") {
					const m = pageMarginsFromAttributes(token.attrs);
					if (m !== undefined) out.pageMargins = m;
				} else if (name === "pageSetup") {
					const ps = pageSetupFromAttributes(token.attrs);
					if (ps !== undefined) out.pageSetup = ps;
				} else if (name === "headerFooter") {
					hf = headerFooterFromAttributes(token.attrs);
					if (token.selfClosing) {
						if (Object.keys(hf).length > 0) out.headerFooter = hf;
						hf = undefined;
					} else {
						inHeaderFooter = true;
					}
				}
			} else if (inHeaderFooter && depth === 2 && isHeaderFooterChild(name)) {
				hfChild = name;
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
 * `location`, a `tooltip`, and a `display` override. Without a target, `location` is an
 * in-workbook reference; with a target, it is a fragment within that external destination. The
 * external target is joined from `rels` here so callers get the resolved URL, not a bare
 * relationship id. A missing/empty `ref`, or a link left with no destination after relationship
 * resolution and location decoding, is skipped.
 */
export function parseHyperlinks(xml: string, rels?: Map<string, Relationship>): Hyperlink[] {
	const links: Hyperlink[] = [];
	for (const token of tokenize(xml)) {
		if (token.kind !== "open" || localName(token.name) !== "hyperlink") continue;
		const ref = token.attrs.ref;
		if (ref === undefined || ref === "") continue;

		// Built immutably (Hyperlink is readonly): include only the attributes that are present.
		const rid = relationshipId(token.attrs);
		const resolvedTarget = rid !== undefined ? rels?.get(rid)?.target : undefined;
		const target =
			resolvedTarget !== undefined && resolvedTarget !== "" ? resolvedTarget : undefined;
		// XML entities are already decoded. Decode the complete ST_Xstring once; omit a result the
		// public writer cannot safely emit (for example, `_x0001_`), then drop a link left with no
		// destination.
		const rawLocation = token.attrs.location;
		const decodedLocation = rawLocation !== undefined ? decodeXstring(rawLocation) : undefined;
		const location =
			decodedLocation !== undefined && decodedLocation !== "" && isXmlSafe(decodedLocation)
				? decodedLocation
				: undefined;
		if (target === undefined && location === undefined) continue;
		const tooltip = token.attrs.tooltip;
		const display = token.attrs.display;
		links.push({
			ref,
			...(target !== undefined ? { target } : {}),
			...(location !== undefined ? { location } : {}),
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

/**
 * Column width/visibility declarations from `<cols>`, in document order. Entries carrying only a
 * default STYLE (no width, not hidden) are style plumbing, not geometry — they are omitted here
 * (style resolution already honors them). Ranges outside Excel's grid and out-of-ceiling widths
 * degrade to absent. Only a `<col>` the shared structure tracker admits, a direct child of the
 * worksheet's canonical `<cols>`, is a column (F11.38): a `<col>` in an extension, in `sheetPr`, or
 * in any other container supplies nothing.
 *
 * @example parseColumnProps('<worksheet><cols><col min="1" max="1" width="12"/></cols><sheetData/></worksheet>') // [{ min: 1, max: 1, width: 12 }]
 * @example parseColumnProps('<worksheet><sheetPr><extLst><ext><col min="1" max="1" width="77"/></ext></extLst></sheetPr><sheetData/></worksheet>') // []
 */
export function parseColumnProps(xml: string): ColumnProps[] {
	const out: ColumnProps[] = [];
	const structure = createWorksheetStructure();
	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		if (token.kind === "close") {
			pushWorksheetClose(structure, token.name);
			continue;
		}
		// <cols> always precedes <sheetData>; stop scanning at the worksheet's own cell region, as
		// the scan always has. Only the canonical one counts, not a sheetData-named extension element.
		const direct = structure.phase === "worksheet" && structure.directChild === undefined;
		if (direct && localName(token.name) === "sheetData") break;
		if (pushWorksheetOpen(structure, token.name, token.selfClosing) !== "col") continue;
		const min = Number(token.attrs.min);
		const max = Number(token.attrs.max);
		if (!Number.isInteger(min) || !Number.isInteger(max)) continue;
		if (min < 1 || max < min || max > MAX_COL) continue;
		const props: { min: number; max: number; width?: number; hidden?: boolean } = { min, max };
		if (token.attrs.width !== undefined) {
			const width = Number(token.attrs.width);
			if (Number.isFinite(width) && width > 0 && width <= MAX_COL_WIDTH) props.width = width;
		}
		if (trueAttribute(token.attrs.hidden)) props.hidden = true;
		if (props.width !== undefined || props.hidden) out.push(props);
	}
	return out;
}

/**
 * Per-row height/visibility from `<row ht hidden>` attributes, keyed by 1-based row index (the
 * `r` attribute, or positional — one past the previous row — when absent, matching the row
 * assembler). Rows with neither property are absent; heights outside (0, 409.5] degrade. Only rows
 * the shared structure tracker admits count, for both the properties and the positional index
 * (F11.38): a `<row>` in an extension or inside a cell supplies nothing and moves no later row.
 *
 * @example parseRowProperties('<worksheet><sheetData><row r="2" ht="20" customHeight="1"/></sheetData></worksheet>') // Map { 2 => { height: 20 } }
 * @example parseRowProperties('<worksheet><sheetData><row r="2"><c r="A2"><extLst><ext><row r="2" ht="300"/></ext></extLst></c></row></sheetData></worksheet>') // Map {}
 */
export function parseRowProperties(xml: string): Map<number, RowProps> {
	const out = new Map<number, RowProps>();
	const structure = createWorksheetStructure();
	let lastRow = 0;
	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		if (token.kind === "close") {
			pushWorksheetClose(structure, token.name);
			continue;
		}
		if (pushWorksheetOpen(structure, token.name, token.selfClosing) !== "row") continue;
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
		if (trueAttribute(token.attrs.hidden)) props.hidden = true;
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
