import { XlsxError } from "../errors";
import {
	type DecodeContext,
	type DefinedName,
	mimeForMediaPath,
	parseConditionalFormatting,
	parseDataValidations,
	parseDrawing,
	parseDxfs,
	parseRels,
	parseSharedStrings,
	parseStyles,
	parseTable,
	parseTheme,
	parseWorkbook,
	type Relationship,
	resolveColor as resolveColorAgainst,
	resolveTarget,
	type StyleTable,
	type ThemeColors,
} from "../ooxml";
import type {
	Cell,
	CellStyle,
	Color,
	ColumnProps,
	Comment,
	ConditionalFormatting,
	DataValidation,
	DxfStyle,
	FreezePane,
	HeaderFooter,
	Hyperlink,
	PageMargins,
	PageSetup,
	PrintOptions,
	RowProps,
	SheetAutoFilter,
	SheetImage,
	SheetInfo,
	SheetProtection,
	SheetState,
	TableInfo,
	WorkbookProtection,
	Worksheet,
} from "../types";
import { openZip, type ZipArchive } from "../zip";
import { decodeText } from "./decode";
import { directoryOf, relsPathFor } from "./opc-path";
import { claimPart } from "./part-claims";
import {
	type PrintSetup,
	parseAutoFilter,
	parseCellStyles,
	parseColumnProps,
	parseComments,
	parseDimension,
	parseFormulas,
	parseFreezePane,
	parseHyperlinks,
	parseMergedCells,
	parsePrintSetup,
	parseRowProperties,
	parseSheetProtection,
	type Row,
	readRows,
	streamRows,
} from "./worksheet";

// The reader's public entry point. `openXlsx` follows the OPC relationship graph — never
// guessed filenames — from the package root to the workbook, then to each worksheet and the
// shared string table, and returns a Workbook of typed cells.
//
// Worksheet XML is decompressed up front (so cell access is synchronous) but only parsed
// into cells on first use; a sheet you never touch costs a decompression, not a parse.

// Relationship type URIs end in these segments; matching the suffix avoids hard-coding the
// 2006 namespace and tolerates the strict/transitional variants.
const REL_OFFICE_DOCUMENT = "/officeDocument";
const REL_SHARED_STRINGS = "/sharedStrings";
const REL_STYLES = "/styles";
const REL_COMMENTS = "/comments";
const REL_THEME = "/theme";
// The drawingML rel; note this does NOT match "/vmlDrawing" (the char before "drawing" there is a
// letter, not the "/"), so the legacy comment drawing is never mistaken for a picture drawing.
const REL_DRAWING = "/drawing";
const REL_TABLE = "/table";

async function readText(zip: ZipArchive, path: string): Promise<string> {
	if (!zip.has(path))
		throw new XlsxError("missing-part", `xlsx is missing a required part: ${path}`);
	return decodeText(await zip.read(path));
}

/** A drawing part a sheet owns (F11.96), with its rels part when the sheet owns that too. */
export interface OwnedDrawing {
	readonly path: string;
	/** The drawing's rels part, or `undefined` when it is missing or another role claimed it first. */
	readonly rels: string | undefined;
}

/**
 * The drawing parts a sheet owns (F11.96): each drawing relationship's part, in relationship order,
 * claimed in this open's set together with the drawing's rels part. The first sheet in tab order
 * that names a drawing part owns it, and a repeat relationship, on the same sheet or a later one, is
 * skipped. Nothing is read.
 *
 * @example claimDrawings(zip, "xl/worksheets/sheet1.xml", rels, claimed) // [{ path: "xl/drawings/drawing1.xml", rels: "xl/drawings/_rels/drawing1.xml.rels" }]
 * @example claimDrawings(zip, "xl/worksheets/sheet2.xml", relsNamingDrawing1, claimed) // [] after sheet 1 claimed it
 * @example claimDrawings(zip, "xl/worksheets/sheet1.xml", undefined, claimed) // []: no rels part
 */
export function claimDrawings(
	zip: ZipArchive,
	sheetPath: string,
	sheetRels: ReadonlyMap<string, Relationship> | undefined,
	claimed: Set<string>,
): OwnedDrawing[] {
	if (sheetRels === undefined) return [];
	const owned: OwnedDrawing[] = [];
	for (const rel of sheetRels.values()) {
		if (!rel.type.endsWith(REL_DRAWING) || rel.targetMode === "External") continue;
		const path = resolveTarget(directoryOf(sheetPath), rel.target);
		if (!zip.has(path) || !claimPart(claimed, path)) continue;
		const relsPath = relsPathFor(path);
		owned.push({
			path,
			rels: zip.has(relsPath) && claimPart(claimed, relsPath) ? relsPath : undefined,
		});
	}
	return owned;
}

/**
 * The bytes of the media part at `path`, read once per workbook: the first call starts the read and
 * caches its promise in `cache`, and later calls (concurrent ones too) share it.
 *
 * @example await cachedRead(zip, cache, "xl/media/image1.png") // reads the part
 * @example await cachedRead(zip, cache, "xl/media/image1.png") // the same Uint8Array, not read again
 */
export function cachedRead(
	zip: ZipArchive,
	cache: Map<string, Promise<Uint8Array>>,
	path: string,
): Promise<Uint8Array> {
	const cached = cache.get(path);
	if (cached !== undefined) return cached;
	const read = zip.read(path);
	cache.set(path, read);
	return read;
}

/**
 * Collect a sheet's pictures (F6.2) from the drawing parts it owns (see {@link claimDrawings}). It
 * follows the chain the file uses: the drawing lists the pictures, and each picture points to an image
 * file. Anything broken or missing along the way (no image, an odd anchor) just skips that one picture
 * instead of failing. Each owned drawing and its rels part are read here, once; image files are read
 * through `media`, which the workbook's sheets share, so pictures using one image file share one
 * `bytes` buffer (F11.96). An image path this open claimed in another role (`claimed`: a part read at
 * open, or a drawing or drawing-rels part) is not read as an image.
 *
 * @example await loadSheetImages(zip, [], claimed, media) // []: the sheet owns no drawing
 * @example await loadSheetImages(zip, [{ path: "xl/drawings/drawing1.xml", rels: "xl/drawings/_rels/drawing1.xml.rels" }], claimed, media) // [{ anchor, bytes, mime }]
 */
export async function loadSheetImages(
	zip: ZipArchive,
	drawings: readonly OwnedDrawing[],
	claimed: ReadonlySet<string>,
	media: Map<string, Promise<Uint8Array>>,
): Promise<readonly SheetImage[]> {
	const images: SheetImage[] = [];
	for (const drawing of drawings) {
		const anchors = parseDrawing(decodeText(await zip.read(drawing.path)));
		if (anchors.length === 0 || drawing.rels === undefined) continue;

		// The drawing's rels map each picture's r:embed to its media part.
		const embedRels = parseRels(decodeText(await zip.read(drawing.rels)));
		const drawingDir = directoryOf(drawing.path);

		for (const anchor of anchors) {
			const rel = embedRels.get(anchor.embed);
			if (rel === undefined || rel.targetMode === "External") continue;
			const mediaPath = resolveTarget(drawingDir, rel.target);
			if (!zip.has(mediaPath) || claimed.has(mediaPath)) continue;
			const bytes = await cachedRead(zip, media, mediaPath);
			images.push({
				anchor: {
					from: anchor.from,
					...(anchor.to !== undefined ? { to: anchor.to } : {}),
					...(anchor.ext !== undefined ? { ext: anchor.ext } : {}),
					...(anchor.editAs !== undefined ? { editAs: anchor.editAs } : {}),
				},
				bytes,
				mime: mimeForMediaPath(mediaPath),
				...(anchor.name !== undefined ? { name: anchor.name } : {}),
			});
		}
	}
	return images;
}

// The xlsx-backed implementation of the public {@link Worksheet} interface (M7 renamed this from
// `Worksheet` to make room for the multi-format seam — the public type is now the interface in
// ../types, which this and the ODS reader's `OdsWorksheet` both satisfy). Bodies are unchanged.
// Exported for white-box tests only; NOT re-exported from the package index, so not public API.
export class XlsxWorksheet implements Worksheet {
	/** Sheet name as shown on Excel's tab. */
	readonly name: string;
	readonly #info: SheetInfo;
	readonly #xml: string;
	readonly #context: DecodeContext;
	readonly #rels: Map<string, Relationship> | undefined;
	readonly #commentsXml: string | undefined;
	readonly #tablesXml: readonly string[];
	// Lazy picture loader (F6.2): reading media bytes needs async decompression, so images are
	// resolved on first `images()` call, not at open — a sheet whose pictures you never touch costs
	// nothing. The promise is cached so concurrent calls share one resolution.
	readonly #loadImages: (() => Promise<readonly SheetImage[]>) | undefined;
	#imagesPromise: Promise<readonly SheetImage[]> | undefined;

	#cells: Map<string, Cell> | undefined;
	#merged: readonly string[] | undefined;
	#hyperlinks: readonly Hyperlink[] | undefined;
	#formulas: Map<string, string> | undefined;
	#cellStyles: Map<string, number> | undefined;
	#dimension: string | undefined;
	#dimensionRead = false;
	#comments: readonly Comment[] | undefined;
	#tables: readonly TableInfo[] | undefined;
	#dataValidations: readonly DataValidation[] | undefined;
	#conditionalFormatting: readonly ConditionalFormatting[] | undefined;
	#columns: readonly ColumnProps[] | undefined;
	#rowProps: ReadonlyMap<number, RowProps> | undefined;
	#freeze: FreezePane | undefined;
	#freezeRead = false;
	#autoFilter: SheetAutoFilter | undefined;
	#autoFilterRead = false;
	#protection: SheetProtection | undefined;
	#protectionRead = false;
	#printSetup: PrintSetup | undefined;

	constructor(
		info: SheetInfo,
		xml: string,
		context: DecodeContext,
		rels?: Map<string, Relationship>,
		commentsXml?: string,
		loadImages?: () => Promise<readonly SheetImage[]>,
		tablesXml?: readonly string[],
	) {
		this.name = info.name;
		this.#info = info;
		this.#xml = xml;
		this.#context = context;
		this.#rels = rels;
		this.#commentsXml = commentsXml;
		this.#loadImages = loadImages;
		this.#tablesXml = tablesXml ?? [];
	}

	/** Workbook-relative part path, e.g. `xl/worksheets/sheet1.xml`. */
	get path(): string {
		return this.#info.path;
	}

	/** false for hidden or very-hidden sheets. */
	get visible(): boolean {
		return this.#info.visible;
	}

	/** The tab's visibility state (F4.6): `"visible"`, `"hidden"`, or `"veryHidden"`. */
	get state(): SheetState {
		return this.#info.state;
	}

	/**
	 * Merged-cell ranges in A1 notation (e.g. `['A1:B1', 'A2:A4']`), in document order. Only the
	 * top-left cell of a merge holds a value; the rest read as `empty`. Empty when none.
	 */
	get mergedCells(): readonly string[] {
		if (this.#merged === undefined) this.#merged = parseMergedCells(this.#xml);
		return this.#merged;
	}

	/**
	 * Hyperlinks declared on this sheet, in document order. Each carries the covered `ref` and,
	 * where present, a resolved external `target`, a `location`, a `tooltip`, and a `display`
	 * override. Without `target`, `location` is an in-workbook reference; with `target`, it is a
	 * fragment within that external destination. Empty when none.
	 */
	get hyperlinks(): readonly Hyperlink[] {
		if (this.#hyperlinks === undefined) {
			this.#hyperlinks = parseHyperlinks(this.#xml, this.#rels);
		}
		return this.#hyperlinks;
	}

	/**
	 * The number-format code applied to the cell at `ref` — a custom code like `"yyyy-mm-dd"`
	 * or `"0.00%"`, or a built-in one. `undefined` when the workbook has no style table, the
	 * id has no portable code, or the file's code exceeds Excel's 255-character limit (such a
	 * cell never reads as a date). An unstyled or absent cell resolves to the default format
	 * (style 0, usually `"General"`), mirroring how date detection defaults.
	 */
	numberFormat(ref: string): string | undefined {
		return this.#context.styles?.formatCode(this.#cellStyleMap().get(ref));
	}

	/**
	 * The resolved style of the cell at `ref` — number format code, font, fill, border, and
	 * alignment (F4.1). Resolution shares the same effective-style map as {@link numberFormat}
	 * (cell `s` → row `customFormat` default → column default), so the two always agree.
	 * `undefined` for an unstyled cell, an absent cell, or a workbook with no style table —
	 * "no style" and "the default style" are deliberately the same answer. Objects are cached
	 * per distinct format record: two cells sharing a format return the same object.
	 */
	style(ref: string): CellStyle | undefined {
		return this.#context.styles?.cellStyle(this.#cellStyleMap().get(ref));
	}

	/**
	 * The sheet's declared used range in A1 notation (e.g. `"A1:E10"`, or a single cell), from
	 * the worksheet's `<dimension>`. `undefined` when the producer omits it — it is an optional
	 * hint, not authoritative, so treat a present value as advisory.
	 */
	get dimension(): string | undefined {
		if (!this.#dimensionRead) {
			this.#dimension = parseDimension(this.#xml);
			this.#dimensionRead = true;
		}
		return this.#dimension;
	}

	/**
	 * The comments anchored to cells on this sheet, in document order — each with its `ref`,
	 * resolved `author`, and plain `text`. Empty when the sheet has no comments part.
	 */
	get comments(): readonly Comment[] {
		if (this.#comments === undefined) {
			this.#comments =
				this.#commentsXml === undefined ? [] : parseComments(this.#commentsXml);
		}
		return this.#comments;
	}

	/**
	 * The defined tables on this sheet (`xl/tables/tableN.xml`), in document order — each with its
	 * `name`, `ref`, `columns`, header/totals flags, and built-in `style`. Empty when the sheet has
	 * no table parts. Parsed lazily on first access.
	 */
	get tables(): readonly TableInfo[] {
		if (this.#tables === undefined) {
			const parsed: TableInfo[] = [];
			for (const xml of this.#tablesXml) {
				const table = parseTable(xml, this.#context.dxfs);
				if (table !== undefined) parsed.push(table);
			}
			this.#tables = parsed;
		}
		return this.#tables;
	}

	/**
	 * The data-validation rules on this sheet, in document order — each with its covered `sqref`
	 * ranges, `type`, operand `formula1`/`formula2` (verbatim), and prompt/error text. Worksheet-level
	 * x14 validations (Excel 2010+ cross-sheet list sources, under `<extLst>`) are skipped. Empty when
	 * the sheet declares none. Parsed lazily on first access.
	 */
	get dataValidations(): readonly DataValidation[] {
		if (this.#dataValidations === undefined) {
			this.#dataValidations = parseDataValidations(this.#xml);
		}
		return this.#dataValidations;
	}

	/**
	 * The conditional-formatting blocks on this sheet, in document order — each with its covered
	 * `sqref` ranges and its `<cfRule>`s. A highlight rule's look is resolved to an inline
	 * {@link DxfStyle} against the workbook `<dxfs>` table; the numeric `dxfId` never surfaces. x14
	 * extensions (under `<extLst>`) are skipped. Empty when the sheet declares none. Parsed lazily.
	 */
	get conditionalFormatting(): readonly ConditionalFormatting[] {
		if (this.#conditionalFormatting === undefined) {
			this.#conditionalFormatting = parseConditionalFormatting(this.#xml, this.#context.dxfs);
		}
		return this.#conditionalFormatting;
	}

	/**
	 * The sheet's autoFilter range (filter dropdowns), or `undefined` when the sheet declares none
	 * (F10.2). Kept SYMBOLIC (the A1 range, never expanded per-cell); per-column filter criteria and
	 * sort state are not modelled. Excel's paired hidden `_xlnm._FilterDatabase` defined name is folded
	 * into this field and stripped from {@link Workbook.definedNames}, so a filter is represented once.
	 * Parsed lazily on first access.
	 */
	get autoFilter(): SheetAutoFilter | undefined {
		if (!this.#autoFilterRead) {
			this.#autoFilter = parseAutoFilter(this.#xml);
			this.#autoFilterRead = true;
		}
		return this.#autoFilter;
	}

	/**
	 * The sheet's `<sheetProtection>` (F10.3), or `undefined` when it declares none. Boolean flags and
	 * any password material are carried verbatim (never computed/verified). Parsed lazily on first access.
	 */
	get protection(): SheetProtection | undefined {
		if (!this.#protectionRead) {
			this.#protection = parseSheetProtection(this.#xml);
			this.#protectionRead = true;
		}
		return this.#protection;
	}

	// The four print-setup elements share ONE token pass (parsePrintSetup returns an object, never
	// undefined, so its presence in the cache doubles as the "already parsed" flag). Lazy on first access.
	#printSetupOf(): PrintSetup {
		if (this.#printSetup === undefined) this.#printSetup = parsePrintSetup(this.#xml);
		return this.#printSetup;
	}

	/** The sheet's page margins (F10.4), or `undefined` when it declares none. Parsed lazily. */
	get pageMargins(): PageMargins | undefined {
		return this.#printSetupOf().pageMargins;
	}

	/** The sheet's page setup (F10.4), or `undefined` when it declares none. Parsed lazily. */
	get pageSetup(): PageSetup | undefined {
		return this.#printSetupOf().pageSetup;
	}

	/** The sheet's print options (F10.4), or `undefined` when it declares none. Parsed lazily. */
	get printOptions(): PrintOptions | undefined {
		return this.#printSetupOf().printOptions;
	}

	/** The sheet's header/footer (F10.4), or `undefined` when it declares none. Parsed lazily. */
	get headerFooter(): HeaderFooter | undefined {
		return this.#printSetupOf().headerFooter;
	}

	/**
	 * The formula text of the cell at an A1 reference, or `undefined` when the cell has none (F5.4).
	 * The text is the stored form (no leading `=`). Shared-formula dependents return the master's
	 * text translated to their position; array formulas return the master's text. `Cell.value` still
	 * holds the cached result, so a formula and its last computed value are read independently.
	 */
	formula(ref: string): string | undefined {
		if (this.#formulas === undefined) this.#formulas = parseFormulas(this.#xml);
		return this.#formulas.get(ref);
	}

	/**
	 * Column width/visibility declarations (`<cols>`), in document order (F4.5). Entries carrying
	 * only a column-default STYLE are not geometry and are omitted — `style(ref)` already resolves
	 * them. Empty when the sheet declares none.
	 */
	get columns(): readonly ColumnProps[] {
		if (this.#columns === undefined) this.#columns = parseColumnProps(this.#xml);
		return this.#columns;
	}

	/**
	 * Per-row height/visibility, keyed by 1-based row index (F4.5). Only rows that declare a
	 * height or hidden flag appear. Empty map when none do.
	 */
	get rowProperties(): ReadonlyMap<number, RowProps> {
		if (this.#rowProps === undefined) this.#rowProps = parseRowProperties(this.#xml);
		return this.#rowProps;
	}

	/**
	 * The sheet's frozen pane, or `undefined` when nothing is frozen (F4.5). Split (non-frozen)
	 * panes are not modelled and read as `undefined`.
	 */
	get freeze(): FreezePane | undefined {
		if (!this.#freezeRead) {
			this.#freeze = parseFreezePane(this.#xml);
			this.#freezeRead = true;
		}
		return this.#freeze;
	}

	/**
	 * The pictures on this sheet, in order (F6.2). This one is an async method (not a plain property)
	 * because reading image data means unzipping it — so the work is done the first time you call it,
	 * and remembered after. Each {@link SheetImage} has the raw `bytes`, its `mime` type, where it
	 * sits (`anchor`), and an optional `name`; pictures using the same image file share one `bytes`
	 * buffer. Empty when the sheet has no pictures. Shapes, charts, and free-floating (absolute)
	 * pictures aren't included, and a picture whose image file is missing is skipped.
	 *
	 * Media is cached per WORKBOOK (F11.96): each drawing and image file is read at most once, however
	 * many sheets or calls ask for it, and pictures on different sheets that use one image file share
	 * one `bytes` buffer, so treat `bytes` as read-only (copy it before mutating or transferring it). A
	 * sheet whose pictures you never touch still costs nothing. A drawing part named by several sheets
	 * belongs to the first sheet in tab order, whichever sheet's `images()` runs first, and a drawing
	 * named twice by one sheet gives its pictures once.
	 */
	images(): Promise<readonly SheetImage[]> {
		if (this.#imagesPromise === undefined) {
			this.#imagesPromise =
				this.#loadImages === undefined ? Promise.resolve([]) : this.#loadImages();
		}
		return this.#imagesPromise;
	}

	#cellStyleMap(): Map<string, number> {
		if (this.#cellStyles === undefined) this.#cellStyles = parseCellStyles(this.#xml);
		return this.#cellStyles;
	}

	/** The cell at an A1 reference. Absent cells read as `empty` (Excel treats them blank). */
	cell(ref: string): Cell {
		return this.#index().get(ref) ?? { ref, type: "empty", value: null };
	}

	/** Stream the populated rows in document order. Sparse: empty rows/cells are absent. */
	async *rows(): AsyncGenerator<Row> {
		for (const row of readRows(this.#xml, this.#context)) {
			yield row;
		}
	}

	#index(): Map<string, Cell> {
		if (this.#cells === undefined) {
			const cells = new Map<string, Cell>();
			for (const row of readRows(this.#xml, this.#context)) {
				for (const cell of row.cells) cells.set(cell.ref, cell);
			}
			this.#cells = cells;
		}
		return this.#cells;
	}
}

export class Workbook {
	/**
	 * Sheets in tab order, as loaded. A declared sheet is left out when its worksheet part is missing,
	 * external or dangling, or when this open already read that part: an earlier loaded sheet's, or a
	 * workbook-level part such as the styles part (xlsx and xlsb). An xlsx sheet whose name repeats an
	 * earlier loaded sheet's is left out too.
	 */
	readonly sheets: readonly SheetInfo[];
	/**
	 * Defined (named) ranges and constants declared in the workbook, in document order. Empty for
	 * formats that don't carry them (ods/xlsb/csv today). The formula evaluator (`openjsxl/formula`)
	 * resolves the constant and simple-range ones; see {@link DefinedName}.
	 */
	readonly definedNames: readonly DefinedName[];
	/** Workbook-level `<workbookProtection>` (F10.3), or `undefined`. Empty for ods/xlsb/csv. */
	readonly protection: WorkbookProtection | undefined;
	/**
	 * Whether this is a MACRO-ENABLED workbook — an `.xlsm`/`.xltm`, sniffed from the workbook's content
	 * type (F10.5). openjsxl reads such files but writes only `.xlsx`: rewriting one through
	 * {@link workbookToInput} → `writeXlsx` DROPS the VBA project (macros). Check this flag to warn a user
	 * before a rewrite discards their macros. Always `false` for `.xlsx` and for ods/xlsb/csv.
	 */
	readonly macroEnabled: boolean;
	/**
	 * The workbook's date system (F11.10): `true` when `<workbookPr date1904>` selects the legacy
	 * Mac 1904 epoch (serial 0 = 1904-01-01), `false` for the default 1900 system. The reader already
	 * honours it when typing date cells; the formula evaluator uses it for every serial ↔ date
	 * conversion. Always `false` for `.ods`/`.csv`; `.xlsb` reports its `BrtWbProp` `f1904` flag
	 * (absent or truncated → `false`).
	 */
	readonly date1904: boolean;
	readonly #byName: Map<string, Worksheet>;
	readonly #themeXml: string | undefined;
	// Note: `Worksheet` here is the public interface (../types), so this same `Workbook` class is
	// reused by every format's reader — the ODS reader builds `OdsWorksheet` instances and hands
	// them to this constructor unchanged (F7.1).
	// The parsed theme is computed on first resolveColor; `undefined` is a valid result (no theme
	// part, or an unparseable one), so a separate flag records that the parse already ran.
	#theme: ThemeColors | undefined;
	#themeParsed = false;

	constructor(
		sheets: SheetInfo[],
		byName: Map<string, Worksheet>,
		themeXml?: string,
		definedNames: readonly DefinedName[] = [],
		protection?: WorkbookProtection,
		macroEnabled = false,
		date1904 = false,
	) {
		this.sheets = sheets;
		this.#byName = byName;
		this.#themeXml = themeXml;
		this.definedNames = definedNames;
		this.protection = protection;
		this.macroEnabled = macroEnabled;
		this.date1904 = date1904;
	}

	/** The worksheet with this tab name. Throws if there is none. */
	sheet(name: string): Worksheet {
		const worksheet = this.#byName.get(name);
		if (worksheet === undefined) {
			const available = this.sheets.map((s) => s.name).join(", ");
			throw new XlsxError(
				"no-such-sheet",
				`no sheet named ${JSON.stringify(name)}; available: ${available}`,
			);
		}
		return worksheet;
	}

	/**
	 * The raw `xl/theme/theme1.xml`, or `undefined` when the workbook has no theme part (a
	 * present-but-empty part is treated as absent). This is the opaque source the bridge carries so a
	 * rewrite keeps custom theme colors; consumers wanting a concrete color should use
	 * {@link resolveColor} instead of parsing this themselves.
	 */
	get themeXml(): string | undefined {
		return this.#themeXml;
	}

	/**
	 * Resolve a raw {@link Color} — as returned by `Worksheet.style(ref)` — to an 8-digit ARGB
	 * string (F5.3). `rgb` colors normalize to `AARRGGBB`; `{theme, tint?}` colors resolve against
	 * this workbook's theme using Excel's tint algorithm. Returns `undefined` when it can't be
	 * resolved: an `{auto}` color, an `{indexed}` palette color, or a theme color with no theme part
	 * or an out-of-range index.
	 */
	resolveColor(color: Color): string | undefined {
		if (!this.#themeParsed) {
			this.#theme = this.#themeXml === undefined ? undefined : parseTheme(this.#themeXml);
			this.#themeParsed = true;
		}
		return resolveColorAgainst(color, this.#theme);
	}
}

interface LoadedWorkbook {
	readonly zip: ZipArchive;
	/** The part paths this open has read so far (F11.95); later loads claim theirs through it. */
	readonly claimed: Set<string>;
	/** Shared decode context (shared strings, styles, date system) for every sheet. */
	readonly context: DecodeContext;
	/** Sheets in tab order, each with its resolved part path. */
	readonly sheets: ReadonlyArray<{ readonly info: SheetInfo; readonly path: string }>;
	/** Raw `xl/theme/theme1.xml`, when present — for color resolution and theme carry (F5.3). */
	readonly themeXml: string | undefined;
	/** Defined (named) ranges/constants from `<definedNames>`, in document order. */
	readonly definedNames: readonly DefinedName[];
	/** Workbook `<workbookProtection>` (F10.3), or `undefined`. */
	readonly protection: WorkbookProtection | undefined;
	/** Whether the workbook part's content type is macro-enabled (`.xlsm`/`.xltm`) — F10.5. */
	readonly macroEnabled: boolean;
	/** The `<workbookPr date1904>` flag (F11.10) — also inside `context` for cell typing. */
	readonly date1904: boolean;
}

// Read the small parts every sheet depends on — relationships, the workbook, shared strings,
// styles — and resolve each sheet's part path through the relationship graph. Worksheets
// themselves are NOT read here, so this stays cheap whether the caller wants random access
// (openXlsx) or a constant-memory stream (streamSheetRows).
async function loadWorkbook(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<LoadedWorkbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
	const zip = openZip(bytes, options);
	// Every part path this open reads (F11.95). The fixed package parts below are recorded as they are
	// read; the optional shared-strings, styles and theme parts, and every sheet, sheet-rels, comments
	// and table load, skip a path already in the set (the part reads as absent).
	const claimed = new Set<string>();

	// Package relationships → the workbook part.
	const packageRels = parseRels(await readText(zip, "_rels/.rels"));
	claimed.add("_rels/.rels");
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT));
	if (office === undefined) {
		throw new XlsxError("not-xlsx", "not an xlsx: no officeDocument relationship");
	}
	const workbookPath = resolveTarget("", office.target);
	const workbookDir = directoryOf(workbookPath);

	// Macro-enabled sniff (F10.5): the workbook part's content type in [Content_Types].xml is
	// `…(sheet|template).macroEnabled.main+xml` for an `.xlsm`/`.xltm`. openjsxl reads these but writes
	// only `.xlsx`, so a rewrite drops the VBA project — this flag lets a caller warn first. A malformed
	// package with no `[Content_Types].xml` reads as not macro-enabled.
	const macroEnabled =
		zip.has("[Content_Types].xml") &&
		(await readText(zip, "[Content_Types].xml")).includes("macroEnabled.main");
	if (zip.has("[Content_Types].xml")) claimed.add("[Content_Types].xml");

	// Workbook sheet list + date system + the workbook's own relationships.
	const {
		sheets: workbookSheets,
		date1904,
		definedNames,
		protection,
	} = parseWorkbook(await readText(zip, workbookPath));
	claimed.add(workbookPath);
	const workbookRels = parseRels(await readText(zip, relsPathFor(workbookPath)));
	claimed.add(relsPathFor(workbookPath));

	// Shared string table (optional — a workbook may use only inline strings).
	let sharedStrings: string[] = [];
	const sst = [...workbookRels.values()].find((r) => r.type.endsWith(REL_SHARED_STRINGS));
	if (sst !== undefined && sst.targetMode !== "External") {
		const sstPath = resolveTarget(workbookDir, sst.target);
		if (zip.has(sstPath) && claimPart(claimed, sstPath)) {
			sharedStrings = parseSharedStrings(decodeText(await zip.read(sstPath)));
		}
	}

	// Style table (optional) — needed to tell date-styled numbers from plain ones. The same part
	// also holds `<dxfs>` (F9.3), the differential styles conditional formatting resolves its
	// `dxfId`s against — parsed from the one decoded string so styles.xml is read once.
	let styles: StyleTable | undefined;
	let dxfs: readonly DxfStyle[] = [];
	const stylesRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_STYLES));
	if (stylesRel !== undefined && stylesRel.targetMode !== "External") {
		const stylesPath = resolveTarget(workbookDir, stylesRel.target);
		if (zip.has(stylesPath) && claimPart(claimed, stylesPath)) {
			const stylesText = decodeText(await zip.read(stylesPath));
			styles = parseStyles(stylesText);
			dxfs = parseDxfs(stylesText);
		}
	}
	const context: DecodeContext = {
		sharedStrings,
		date1904,
		...(styles !== undefined ? { styles } : {}),
		...(dxfs.length > 0 ? { dxfs } : {}),
	};

	// Theme part (optional) — the color scheme resolveColor needs and the bytes the bridge carries.
	let themeXml: string | undefined;
	const themeRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_THEME));
	if (themeRel !== undefined && themeRel.targetMode !== "External") {
		const themePath = resolveTarget(workbookDir, themeRel.target);
		if (zip.has(themePath) && claimPart(claimed, themePath)) {
			const decoded = decodeText(await zip.read(themePath));
			// A present-but-EMPTY theme part (a truncated/corrupt producer) is no usable theme —
			// treat it as absent so resolveColor degrades to `undefined` and, crucially, the bridge
			// doesn't carry "" into the writer's non-empty check (which would reject the workbook's
			// own read-back). An empty part and a missing part are semantically identical.
			if (decoded.length > 0) themeXml = decoded;
		}
	}

	// Resolve each sheet's r:id to a part path. A sheet whose worksheet part is missing, external, or
	// dangling is DROPPED — so the loaded list can be shorter than workbook.xml declared, and the
	// survivors compact to a new 0-based order. So is a sheet whose part this open already read (an
	// earlier loaded sheet's, or a workbook-level part): a package that names one worksheet part many
	// times would otherwise decode and keep it once per sheet (F11.95). Track raw→loaded so
	// sheet-scoped defined names can be re-indexed against the sheets that actually loaded.
	const sheets: Array<{ info: SheetInfo; path: string }> = [];
	const rawToLoaded = new Map<number, number>();
	const seenNames = new Set<string>();
	for (let raw = 0; raw < workbookSheets.length; raw++) {
		const entry = workbookSheets[raw];
		if (entry === undefined) continue;
		const rel = workbookRels.get(entry.rid);
		if (rel === undefined || rel.targetMode === "External") continue;
		const path = resolveTarget(workbookDir, rel.target);
		if (!zip.has(path)) continue;
		// parseWorkbook has already decoded CT_Sheet's ST_Xstring name. A malformed producer can
		// therefore declare distinct wire spellings that collapse to the same logical (or same-case-
		// folded) tab name, e.g. `A` and `_x0041_`. Match the writer's case-insensitive uniqueness
		// rule and the reader's existing first-definition-wins lookup: keep the first READABLE sheet,
		// drop later aliases, and leave them absent from rawToLoaded so scoped names reconcile below.
		const nameKey = entry.name.toLowerCase();
		if (seenNames.has(nameKey)) continue;
		// Only a loaded sheet claims its path, and a sheet dropped for its name claims nothing.
		if (!claimPart(claimed, path)) continue;
		seenNames.add(nameKey);
		rawToLoaded.set(raw, sheets.length);
		sheets.push({
			info: { name: entry.name, path, visible: entry.visible, state: entry.state },
			path,
		});
	}

	// Reconcile sheet-scoped defined names to the LOADED sheets (F10.1). parseWorkbook validated each
	// localSheetId against the RAW workbook.xml sheet count, but dropping a sheet both shrinks and
	// re-indexes that list — so remap a scoped name's localSheetId to its sheet's new position, and DROP
	// a name whose scope sheet didn't load. This keeps every localSheetId a valid index into the loaded
	// `sheets` — the shared-model invariant the writer's validateDefinedNames relies on (it checks
	// localSheetId against the loaded count). Nothing dropped → the map is the identity and the array
	// passes through untouched (byte-identity path).
	const reconciledNames =
		rawToLoaded.size === workbookSheets.length
			? definedNames
			: definedNames.flatMap((dn) => {
					if (dn.localSheetId === undefined) return [dn];
					const loaded = rawToLoaded.get(dn.localSheetId);
					if (loaded === undefined) return [];
					return loaded === dn.localSheetId ? [dn] : [{ ...dn, localSheetId: loaded }];
				});

	return {
		zip,
		claimed,
		context,
		sheets,
		themeXml,
		definedNames: reconciledNames,
		protection,
		macroEnabled,
		date1904,
	};
}

/**
 * Reader options — two zip-bomb guards plus optional ZIP checksum verification.
 *
 * - `maxPartBytes`: absolute ceiling on any one part's decompressed output (default 2 GiB). A part
 *   that declares — or, while inflating, actually produces — more fails typed (`part-too-large`).
 * - `maxCompressionRatio`: a part may not inflate to more than this multiple of its compressed size
 *   (default 300), above an 8 MiB floor. Catches a decompression bomb (deflate reaches ~1000:1)
 *   early, while real spreadsheet XML (measured ~10–33:1, even for the most repetitive sheets) and
 *   small high-ratio parts pass untouched.
 * - `verifyChecksums`: verify each fully consumed ZIP entry against its central-directory CRC-32
 *   (default false). CRC-32 detects accidental corruption; it is not authentication or tamper
 *   resistance. A cancelled stream makes no checksum claim for its partially consumed entry.
 *
 * Both guards apply to the constant-memory streaming path (`streamSheetRows`) too, so a *single*
 * legitimately huge worksheet whose decompressed XML exceeds `maxPartBytes` (2 GiB) is rejected by
 * default even when streamed — pass `maxPartBytes: Number.POSITIVE_INFINITY` to stream it (the ratio
 * guard still protects against a bomb).
 */
export interface ReadOptions {
	readonly maxPartBytes?: number;
	readonly maxCompressionRatio?: number;
	readonly verifyChecksums?: boolean;
}

export async function openXlsx(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<Workbook> {
	const {
		zip,
		claimed,
		context,
		sheets,
		themeXml,
		definedNames,
		protection,
		macroEnabled,
		date1904,
	} = await loadWorkbook(source, options);

	// Decompress each worksheet (so cell access is synchronous) and build the Worksheet. Image parts
	// are read lazily, through one cache the workbook's sheets share (F11.96).
	const pictureReads = new Map<string, Promise<Uint8Array>>();
	const infos: SheetInfo[] = [];
	const byName = new Map<string, Worksheet>();
	for (const { info, path } of sheets) {
		const xml = decodeText(await zip.read(path));
		// The sheet's own relationships (xl/worksheets/_rels/sheetN.xml.rels) resolve hyperlink
		// r:ids and locate the comments part. Optional — a plain sheet has no rels part. This and the
		// comments and table parts below are each read at most once per open (F11.95): a part this
		// open already read is skipped, and the sheet loads without it.
		const relsPath = relsPathFor(path);
		const rels =
			zip.has(relsPath) && claimPart(claimed, relsPath)
				? parseRels(decodeText(await zip.read(relsPath)))
				: undefined;

		// Comments live in a separate part linked from the worksheet rels.
		let commentsXml: string | undefined;
		const commentsRel = rels && [...rels.values()].find((r) => r.type.endsWith(REL_COMMENTS));
		if (commentsRel !== undefined && commentsRel.targetMode !== "External") {
			const commentsPath = resolveTarget(directoryOf(path), commentsRel.target);
			if (zip.has(commentsPath) && claimPart(claimed, commentsPath)) {
				commentsXml = decodeText(await zip.read(commentsPath));
			}
		}

		// Tables are one part each (a sheet may own several), linked from the worksheet rels.
		const tablesXml: string[] = [];
		if (rels !== undefined) {
			for (const rel of rels.values()) {
				if (!rel.type.endsWith(REL_TABLE) || rel.targetMode === "External") continue;
				const tablePath = resolveTarget(directoryOf(path), rel.target);
				if (zip.has(tablePath) && claimPart(claimed, tablePath)) {
					tablesXml.push(decodeText(await zip.read(tablePath)));
				}
			}
		}

		infos.push(info);
		// First definition wins if two sheets somehow share a name.
		if (!byName.has(info.name)) {
			// A lazy picture loader, invoked only on the first Worksheet.images() call so a sheet whose
			// images you never read costs no media I/O. The drawing parts it may read are claimed now,
			// in tab order, so which sheet owns a drawing never depends on call order (F11.96).
			const drawings = claimDrawings(zip, path, rels, claimed);
			const loadImages = (): Promise<readonly SheetImage[]> =>
				loadSheetImages(zip, drawings, claimed, pictureReads);
			byName.set(
				info.name,
				new XlsxWorksheet(info, xml, context, rels, commentsXml, loadImages, tablesXml),
			);
		}
	}

	return new Workbook(infos, byName, themeXml, definedNames, protection, macroEnabled, date1904);
}

/**
 * Stream the rows of one sheet with roughly constant memory: the worksheet is never
 * materialized as a whole string — it is decompressed and tokenized chunk by chunk, and each
 * row is yielded then discarded (F2.2). Use this for large sheets; use `openXlsx` when you
 * need random `cell()` access. `sheetName` defaults to the first sheet in tab order.
 */
export async function* streamSheetRows(
	source: Uint8Array | ArrayBuffer,
	sheetName?: string,
	options?: ReadOptions,
): AsyncGenerator<Row> {
	const { zip, context, sheets } = await loadWorkbook(source, options);
	const first = sheets[0];
	if (first === undefined) throw new XlsxError("not-xlsx", "xlsx has no readable worksheets");

	let path = first.path;
	if (sheetName !== undefined) {
		const match = sheets.find((s) => s.info.name === sheetName);
		if (match === undefined) {
			const available = sheets.map((s) => s.info.name).join(", ");
			throw new XlsxError(
				"no-such-sheet",
				`no sheet named ${JSON.stringify(sheetName)}; available: ${available}`,
			);
		}
		path = match.path;
	}

	yield* streamRows(zip.readStream(path), context);
}
