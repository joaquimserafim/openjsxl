import { XlsxError } from "../errors";
import { parseRef, parseRels, resolveTarget } from "../ooxml";
import type {
	Cell,
	CellStyle,
	ColumnProps,
	Comment,
	ConditionalFormatting,
	DataValidation,
	FreezePane,
	Hyperlink,
	Row,
	RowProps,
	SheetImage,
	SheetInfo,
	SheetState,
	TableInfo,
	Worksheet,
} from "../types";
import {
	parseXlsbSheet,
	parseXlsbStrings,
	parseXlsbStyles,
	parseXlsbWorkbook,
	type XlsbStyleTable,
} from "../xlsb";
import { openZip, type ZipArchive } from "../zip";
import { type ReadOptions, Workbook } from "./workbook";

// The `.xlsb` reader. An Excel Binary Workbook is the SAME OPC container as `.xlsx` — the package
// relationships and content types are XML, so the F1.4 relationship graph is reused verbatim — but
// the workbook / worksheets / sharedStrings / styles parts are BIFF12 binary (`.bin`). openXlsb
// resolves each part through the rels graph, parses the binary records (see ../biff, ../xlsb), and
// returns the SAME public Workbook `openXlsx` returns. Styles beyond number formats, formula text,
// comments, geometry, images, and merges are not carried and degrade on their accessors (F7.2).

const decoder = new TextDecoder();

const REL_OFFICE_DOCUMENT = "/officeDocument";
const REL_SHARED_STRINGS = "/sharedStrings";
const REL_STYLES = "/styles";
const REL_HYPERLINK = "/hyperlink";

const NO_MERGES: readonly string[] = [];
const NO_COMMENTS: readonly Comment[] = [];
const NO_TABLES: readonly TableInfo[] = [];
const NO_DATA_VALIDATIONS: readonly DataValidation[] = [];
const NO_CONDITIONAL_FORMATTING: readonly ConditionalFormatting[] = [];
const NO_COLUMNS: readonly ColumnProps[] = [];
const NO_IMAGES: readonly SheetImage[] = [];
const NO_ROW_PROPS: ReadonlyMap<number, RowProps> = new Map();

function directoryOf(path: string): string {
	const slash = path.lastIndexOf("/");
	return slash === -1 ? "" : path.slice(0, slash);
}

function relsPathFor(path: string): string {
	const slash = path.lastIndexOf("/");
	const dir = slash === -1 ? "" : path.slice(0, slash);
	const file = slash === -1 ? path : path.slice(slash + 1);
	return dir === "" ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`;
}

async function readText(zip: ZipArchive, path: string): Promise<string> {
	if (!zip.has(path))
		throw new XlsxError("missing-part", `xlsb is missing a required part: ${path}`);
	return decoder.decode(await zip.read(path));
}

// The xlsb-backed Worksheet: a plain data holder built from the parsed sheet part. numberFormat is
// supported (via the style table); everything else the format doesn't carry degrades to empty.
class XlsbWorksheet implements Worksheet {
	readonly name: string;
	readonly #info: SheetInfo;
	readonly #cells: Map<string, Cell>;
	readonly #cellStyles: Map<string, number>;
	readonly #styles: XlsbStyleTable | undefined;
	readonly #hyperlinks: readonly Hyperlink[];
	readonly #dimension: string | undefined;
	#rowsCache: Row[] | undefined;

	constructor(
		info: SheetInfo,
		cells: Map<string, Cell>,
		cellStyles: Map<string, number>,
		styles: XlsbStyleTable | undefined,
		hyperlinks: readonly Hyperlink[],
		dimension: string | undefined,
	) {
		this.name = info.name;
		this.#info = info;
		this.#cells = cells;
		this.#cellStyles = cellStyles;
		this.#styles = styles;
		this.#hyperlinks = hyperlinks;
		this.#dimension = dimension;
	}

	get path(): string {
		return this.#info.path;
	}

	get visible(): boolean {
		return this.#info.visible;
	}

	get state(): SheetState {
		return this.#info.state;
	}

	get hyperlinks(): readonly Hyperlink[] {
		return this.#hyperlinks;
	}

	get dimension(): string | undefined {
		return this.#dimension;
	}

	/** The number-format code applied at `ref` (via the cell's style), or `undefined`. */
	numberFormat(ref: string): string | undefined {
		return this.#styles?.formatCode(this.#cellStyles.get(ref));
	}

	// ── Not carried by .xlsb in F7.2: degrade to the shared empties. ──
	get mergedCells(): readonly string[] {
		return NO_MERGES;
	}

	get comments(): readonly Comment[] {
		return NO_COMMENTS;
	}

	get tables(): readonly TableInfo[] {
		return NO_TABLES;
	}

	get dataValidations(): readonly DataValidation[] {
		return NO_DATA_VALIDATIONS;
	}

	get conditionalFormatting(): readonly ConditionalFormatting[] {
		return NO_CONDITIONAL_FORMATTING;
	}

	get columns(): readonly ColumnProps[] {
		return NO_COLUMNS;
	}

	get rowProperties(): ReadonlyMap<number, RowProps> {
		return NO_ROW_PROPS;
	}

	get freeze(): FreezePane | undefined {
		return undefined;
	}

	style(_ref: string): CellStyle | undefined {
		return undefined;
	}

	formula(_ref: string): string | undefined {
		return undefined;
	}

	images(): Promise<readonly SheetImage[]> {
		return Promise.resolve(NO_IMAGES);
	}

	cell(ref: string): Cell {
		return this.#cells.get(ref) ?? { ref, type: "empty", value: null };
	}

	async *rows(): AsyncGenerator<Row> {
		for (const row of this.#rows()) yield row;
	}

	#rows(): Row[] {
		if (this.#rowsCache === undefined) {
			const entries = [...this.#cells.values()].map((cell) => {
				const { row, col } = parseRef(cell.ref);
				return { row, col, cell };
			});
			entries.sort((a, b) => a.row - b.row || a.col - b.col);
			const rows: Row[] = [];
			let current: { index: number; cells: Cell[] } | undefined;
			for (const entry of entries) {
				if (current === undefined || current.index !== entry.row) {
					current = { index: entry.row, cells: [] };
					rows.push(current);
				}
				current.cells.push(entry.cell);
			}
			this.#rowsCache = rows;
		}
		return this.#rowsCache;
	}
}

/**
 * Open an Excel Binary Workbook (`.xlsb`) and read its cells. Returns the same {@link Workbook} as
 * {@link openXlsx} — sheets, typed cells (with style-driven date detection), number formats,
 * hyperlinks, dimension, and visibility. Styles beyond number formats, formula text, comments,
 * geometry, images, and merged ranges are not carried and degrade on their accessors.
 */
export async function openXlsb(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<Workbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
	const zip = openZip(bytes, options);

	// Package relationships → the workbook part (BIFF12, but the rels are XML).
	const packageRels = parseRels(await readText(zip, "_rels/.rels"));
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT));
	if (office === undefined) {
		throw new XlsxError("not-xlsx", "not a valid xlsb: no officeDocument relationship");
	}
	const workbookPath = resolveTarget("", office.target);
	const workbookDir = directoryOf(workbookPath);

	const { sheets: sheetEntries, date1904 } = parseXlsbWorkbook(await readPart(zip, workbookPath));
	const workbookRels = parseRels(await readText(zip, relsPathFor(workbookPath)));

	// Shared strings + styles (optional parts, resolved through the workbook rels).
	let sharedStrings: string[] = [];
	const sstRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_SHARED_STRINGS));
	if (sstRel !== undefined && sstRel.targetMode !== "External") {
		const p = resolveTarget(workbookDir, sstRel.target);
		if (zip.has(p)) sharedStrings = parseXlsbStrings(await zip.read(p));
	}
	let styles: XlsbStyleTable | undefined;
	const stylesRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_STYLES));
	if (stylesRel !== undefined && stylesRel.targetMode !== "External") {
		const p = resolveTarget(workbookDir, stylesRel.target);
		if (zip.has(p)) styles = parseXlsbStyles(await zip.read(p));
	}

	const infos: SheetInfo[] = [];
	const byName = new Map<string, Worksheet>();
	for (const entry of sheetEntries) {
		const rel = workbookRels.get(entry.relId);
		if (rel === undefined || rel.targetMode === "External") continue;
		const path = resolveTarget(workbookDir, rel.target);
		if (!zip.has(path)) continue;

		const parsed = parseXlsbSheet(await zip.read(path), sharedStrings, styles, date1904);

		// Resolve hyperlink rel ids through the sheet's own rels part.
		const hyperlinks = await resolveHyperlinks(zip, path, parsed.hyperlinks);

		const info: SheetInfo = {
			name: entry.name,
			path,
			visible: entry.state === "visible",
			state: entry.state,
		};
		infos.push(info);
		if (!byName.has(entry.name)) {
			byName.set(
				entry.name,
				new XlsbWorksheet(
					info,
					parsed.cells,
					parsed.cellStyles,
					styles,
					hyperlinks,
					parsed.dimension,
				),
			);
		}
	}

	return new Workbook(infos, byName);
}

async function readPart(zip: ZipArchive, path: string): Promise<Uint8Array> {
	if (!zip.has(path))
		throw new XlsxError("missing-part", `xlsb is missing a required part: ${path}`);
	return zip.read(path);
}

// Turn each sheet hyperlink's rel id into a resolved external target / in-workbook location, using
// the worksheet's own rels part — exactly as the xlsx reader does.
async function resolveHyperlinks(
	zip: ZipArchive,
	sheetPath: string,
	refs: readonly { readonly ref: string; readonly relId: string }[],
): Promise<readonly Hyperlink[]> {
	if (refs.length === 0) return [];
	const relsPath = relsPathFor(sheetPath);
	const rels = zip.has(relsPath) ? parseRels(await readText(zip, relsPath)) : undefined;
	const sheetDir = directoryOf(sheetPath);
	const links: Hyperlink[] = [];
	for (const { ref, relId } of refs) {
		const rel = rels?.get(relId);
		if (rel === undefined || !rel.type.endsWith(REL_HYPERLINK)) {
			links.push({ ref });
			continue;
		}
		if (rel.targetMode === "External") {
			links.push({ ref, target: rel.target });
		} else {
			links.push({ ref, location: resolveTarget(sheetDir, rel.target) });
		}
	}
	return links;
}
