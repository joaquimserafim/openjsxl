import { XlsxError } from "../errors";
import {
	type DecodeContext,
	type DefinedName,
	mimeForMediaPath,
	parseDrawing,
	parseRels,
	parseSharedStrings,
	parseStyles,
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
	FreezePane,
	Hyperlink,
	RowProps,
	SheetImage,
	SheetInfo,
	SheetState,
	Worksheet,
} from "../types";
import { openZip, type ZipArchive } from "../zip";
import {
	parseCellStyles,
	parseColumnProps,
	parseComments,
	parseDimension,
	parseFormulas,
	parseFreezePane,
	parseHyperlinks,
	parseMergedCells,
	parseRowProperties,
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

const decoder = new TextDecoder();

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

function directoryOf(path: string): string {
	const slash = path.lastIndexOf("/");
	return slash === -1 ? "" : path.slice(0, slash);
}

// The relationships for a part live in `<dir>/_rels/<file>.rels`.
function relsPathFor(path: string): string {
	const slash = path.lastIndexOf("/");
	const dir = slash === -1 ? "" : path.slice(0, slash);
	const file = slash === -1 ? path : path.slice(slash + 1);
	return dir === "" ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`;
}

async function readText(zip: ZipArchive, path: string): Promise<string> {
	if (!zip.has(path))
		throw new XlsxError("missing-part", `xlsx is missing a required part: ${path}`);
	return decoder.decode(await zip.read(path));
}

/**
 * Collect a sheet's pictures (F6.2). It follows the chain the file uses: the sheet points to a
 * drawing, the drawing lists the pictures, and each picture points to an image file — this reads
 * that image data. Anything broken or missing along the way (no drawing, no image, an odd anchor)
 * just skips that one picture instead of failing. Each image file is read only once, so two pictures
 * using the same file share the same bytes.
 */
async function loadSheetImages(
	zip: ZipArchive,
	sheetPath: string,
	sheetRels: Map<string, Relationship> | undefined,
): Promise<readonly SheetImage[]> {
	if (sheetRels === undefined) return [];
	const drawingRels = [...sheetRels.values()].filter(
		(r) => r.type.endsWith(REL_DRAWING) && r.targetMode !== "External",
	);
	if (drawingRels.length === 0) return [];

	const sheetDir = directoryOf(sheetPath);
	const images: SheetImage[] = [];
	for (const drawingRel of drawingRels) {
		const drawingPath = resolveTarget(sheetDir, drawingRel.target);
		if (!zip.has(drawingPath)) continue;
		const anchors = parseDrawing(decoder.decode(await zip.read(drawingPath)));
		if (anchors.length === 0) continue;

		// The drawing's rels map each picture's r:embed to its media part.
		const embedRelsPath = relsPathFor(drawingPath);
		if (!zip.has(embedRelsPath)) continue;
		const embedRels = parseRels(decoder.decode(await zip.read(embedRelsPath)));
		const drawingDir = directoryOf(drawingPath);
		const mediaCache = new Map<string, Uint8Array>();

		for (const anchor of anchors) {
			const rel = embedRels.get(anchor.embed);
			if (rel === undefined || rel.targetMode === "External") continue;
			const mediaPath = resolveTarget(drawingDir, rel.target);
			if (!zip.has(mediaPath)) continue;
			let bytes = mediaCache.get(mediaPath);
			if (bytes === undefined) {
				bytes = await zip.read(mediaPath);
				mediaCache.set(mediaPath, bytes);
			}
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
	#columns: readonly ColumnProps[] | undefined;
	#rowProps: ReadonlyMap<number, RowProps> | undefined;
	#freeze: FreezePane | undefined;
	#freezeRead = false;

	constructor(
		info: SheetInfo,
		xml: string,
		context: DecodeContext,
		rels?: Map<string, Relationship>,
		commentsXml?: string,
		loadImages?: () => Promise<readonly SheetImage[]>,
	) {
		this.name = info.name;
		this.#info = info;
		this.#xml = xml;
		this.#context = context;
		this.#rels = rels;
		this.#commentsXml = commentsXml;
		this.#loadImages = loadImages;
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
	 * where present, a resolved external `target`, an in-workbook `location`, a `tooltip`, and a
	 * `display` override. Empty when none.
	 */
	get hyperlinks(): readonly Hyperlink[] {
		if (this.#hyperlinks === undefined) {
			this.#hyperlinks = parseHyperlinks(this.#xml, this.#rels);
		}
		return this.#hyperlinks;
	}

	/**
	 * The number-format code applied to the cell at `ref` — a custom code like `"yyyy-mm-dd"`
	 * or `"0.00%"`, or a built-in one. `undefined` when the workbook has no style table or the
	 * id has no portable code. An unstyled or absent cell resolves to the default format (style
	 * 0, usually `"General"`), mirroring how date detection defaults.
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
	 * Media is cached per SHEET, not per workbook: when several sheets show the same image, each
	 * sheet's first `images()` call decompresses that image once more (sheets stay independent, and
	 * a sheet whose pictures you never touch costs nothing). Rewriting through the bridge is
	 * unaffected — the writer dedupes identical bytes back into one media part.
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
	/** Sheets in tab order. */
	readonly sheets: readonly SheetInfo[];
	/**
	 * Defined (named) ranges and constants declared in the workbook, in document order. Empty for
	 * formats that don't carry them (ods/xlsb/csv today). The formula evaluator (`openjsxl/formula`)
	 * resolves the constant and simple-range ones; see {@link DefinedName}.
	 */
	readonly definedNames: readonly DefinedName[];
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
	) {
		this.sheets = sheets;
		this.#byName = byName;
		this.#themeXml = themeXml;
		this.definedNames = definedNames;
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
	/** Shared decode context (shared strings, styles, date system) for every sheet. */
	readonly context: DecodeContext;
	/** Sheets in tab order, each with its resolved part path. */
	readonly sheets: ReadonlyArray<{ readonly info: SheetInfo; readonly path: string }>;
	/** Raw `xl/theme/theme1.xml`, when present — for color resolution and theme carry (F5.3). */
	readonly themeXml: string | undefined;
	/** Defined (named) ranges/constants from `<definedNames>`, in document order. */
	readonly definedNames: readonly DefinedName[];
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

	// Package relationships → the workbook part.
	const packageRels = parseRels(await readText(zip, "_rels/.rels"));
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT));
	if (office === undefined) {
		throw new XlsxError("not-xlsx", "not an xlsx: no officeDocument relationship");
	}
	const workbookPath = resolveTarget("", office.target);
	const workbookDir = directoryOf(workbookPath);

	// Workbook sheet list + date system + the workbook's own relationships.
	const {
		sheets: workbookSheets,
		date1904,
		definedNames,
	} = parseWorkbook(await readText(zip, workbookPath));
	const workbookRels = parseRels(await readText(zip, relsPathFor(workbookPath)));

	// Shared string table (optional — a workbook may use only inline strings).
	let sharedStrings: string[] = [];
	const sst = [...workbookRels.values()].find((r) => r.type.endsWith(REL_SHARED_STRINGS));
	if (sst !== undefined && sst.targetMode !== "External") {
		const sstPath = resolveTarget(workbookDir, sst.target);
		if (zip.has(sstPath)) {
			sharedStrings = parseSharedStrings(decoder.decode(await zip.read(sstPath)));
		}
	}

	// Style table (optional) — needed to tell date-styled numbers from plain ones.
	let styles: StyleTable | undefined;
	const stylesRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_STYLES));
	if (stylesRel !== undefined && stylesRel.targetMode !== "External") {
		const stylesPath = resolveTarget(workbookDir, stylesRel.target);
		if (zip.has(stylesPath)) {
			styles = parseStyles(decoder.decode(await zip.read(stylesPath)));
		}
	}
	const context: DecodeContext =
		styles !== undefined ? { sharedStrings, date1904, styles } : { sharedStrings, date1904 };

	// Theme part (optional) — the color scheme resolveColor needs and the bytes the bridge carries.
	let themeXml: string | undefined;
	const themeRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_THEME));
	if (themeRel !== undefined && themeRel.targetMode !== "External") {
		const themePath = resolveTarget(workbookDir, themeRel.target);
		if (zip.has(themePath)) {
			const decoded = decoder.decode(await zip.read(themePath));
			// A present-but-EMPTY theme part (a truncated/corrupt producer) is no usable theme —
			// treat it as absent so resolveColor degrades to `undefined` and, crucially, the bridge
			// doesn't carry "" into the writer's non-empty check (which would reject the workbook's
			// own read-back). An empty part and a missing part are semantically identical.
			if (decoded.length > 0) themeXml = decoded;
		}
	}

	// Resolve each sheet's r:id to a part path.
	const sheets: Array<{ info: SheetInfo; path: string }> = [];
	for (const entry of workbookSheets) {
		const rel = workbookRels.get(entry.rid);
		if (rel === undefined || rel.targetMode === "External") continue;
		const path = resolveTarget(workbookDir, rel.target);
		if (!zip.has(path)) continue;
		sheets.push({
			info: { name: entry.name, path, visible: entry.visible, state: entry.state },
			path,
		});
	}

	return { zip, context, sheets, themeXml, definedNames };
}

/**
 * Reader options. `maxPartBytes` caps the declared decompressed size of any single part — a
 * zip-bomb guard independent of the archive's own (untrusted) size fields. Omit for no ceiling.
 */
export interface ReadOptions {
	readonly maxPartBytes?: number;
}

export async function openXlsx(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<Workbook> {
	const { zip, context, sheets, themeXml, definedNames } = await loadWorkbook(source, options);

	// Decompress each worksheet (so cell access is synchronous) and build the Worksheet.
	const infos: SheetInfo[] = [];
	const byName = new Map<string, Worksheet>();
	for (const { info, path } of sheets) {
		const xml = decoder.decode(await zip.read(path));
		// The sheet's own relationships (xl/worksheets/_rels/sheetN.xml.rels) resolve hyperlink
		// r:ids and locate the comments part. Optional — a plain sheet has no rels part.
		const relsPath = relsPathFor(path);
		const rels = zip.has(relsPath)
			? parseRels(decoder.decode(await zip.read(relsPath)))
			: undefined;

		// Comments live in a separate part linked from the worksheet rels.
		let commentsXml: string | undefined;
		const commentsRel = rels && [...rels.values()].find((r) => r.type.endsWith(REL_COMMENTS));
		if (commentsRel !== undefined && commentsRel.targetMode !== "External") {
			const commentsPath = resolveTarget(directoryOf(path), commentsRel.target);
			if (zip.has(commentsPath)) commentsXml = decoder.decode(await zip.read(commentsPath));
		}

		infos.push(info);
		// First definition wins if two sheets somehow share a name.
		if (!byName.has(info.name)) {
			// A lazy picture loader — captures the zip + this sheet's path/rels, invoked only on the
			// first Worksheet.images() call so a sheet whose images you never read costs no media I/O.
			const loadImages = (): Promise<readonly SheetImage[]> =>
				loadSheetImages(zip, path, rels);
			byName.set(
				info.name,
				new XlsxWorksheet(info, xml, context, rels, commentsXml, loadImages),
			);
		}
	}

	return new Workbook(infos, byName, themeXml, definedNames);
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
