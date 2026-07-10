import { XlsxError } from "../errors";
import { parseOdsContent } from "../ods";
import { parseRef } from "../ooxml";
import type {
	Cell,
	CellStyle,
	ColumnProps,
	Comment,
	FreezePane,
	Hyperlink,
	Row,
	RowProps,
	SheetImage,
	SheetInfo,
	SheetState,
	Worksheet,
} from "../types";
import { openZip } from "../zip";
import { type ReadOptions, Workbook } from "./workbook";

// The `.ods` reader. An OpenDocument spreadsheet is a ZIP whose sheets all live in one
// `content.xml`; there is no OPC relationship graph (sheet order is document order) and cell types
// are explicit, so the whole part is parsed up front into typed cell tables (per-sheet laziness
// can't exist inside a single part). The result is the SAME public Workbook/Worksheet surface
// `openXlsx` returns — accessors ODS can't express (styles, formula text, comments, geometry,
// images) DEGRADE to `[]`/`undefined`, never throw (F7.1).

const decoder = new TextDecoder();

// The ODF spreadsheet media type. A spreadsheet TEMPLATE (`…spreadsheet-template`) shares this
// prefix, so `startsWith` accepts both while still rejecting a text/presentation document.
const MIMETYPE_SPREADSHEET = "application/vnd.oasis.opendocument.spreadsheet";

const NO_COMMENTS: readonly Comment[] = [];
const NO_COLUMNS: readonly ColumnProps[] = [];
const NO_IMAGES: readonly SheetImage[] = [];
const NO_ROW_PROPS: ReadonlyMap<number, RowProps> = new Map();

// The ODS-backed implementation of the shared Worksheet interface. It is a plain data holder: the
// content.xml parse already produced this sheet's cells/merges/hyperlinks/dimension, so every
// accessor is a getter, and the features ODS doesn't carry return the shared empties.
class OdsWorksheet implements Worksheet {
	readonly name: string;
	readonly #info: SheetInfo;
	readonly #cells: Map<string, Cell>;
	readonly #merges: readonly string[];
	readonly #hyperlinks: readonly Hyperlink[];
	readonly #dimension: string | undefined;
	#rowsCache: Row[] | undefined;

	constructor(
		info: SheetInfo,
		cells: Map<string, Cell>,
		merges: readonly string[],
		hyperlinks: readonly Hyperlink[],
		dimension: string | undefined,
	) {
		this.name = info.name;
		this.#info = info;
		this.#cells = cells;
		this.#merges = merges;
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

	get mergedCells(): readonly string[] {
		return this.#merges;
	}

	get hyperlinks(): readonly Hyperlink[] {
		return this.#hyperlinks;
	}

	/** The used range, synthesized from populated cells (ODF stores no `<dimension>`). */
	get dimension(): string | undefined {
		return this.#dimension;
	}

	// ── Accessors ODS does not carry: degrade to the shared empties (F7.1 drop-list). ──
	get comments(): readonly Comment[] {
		return NO_COMMENTS;
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

	numberFormat(_ref: string): string | undefined {
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

	// Group the (already materialized) cells into rows in ascending row/column order, once.
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
 * Open an OpenDocument spreadsheet (`.ods`) and read its cells. Returns the same {@link Workbook}
 * as {@link openXlsx} — sheets, typed cells, merges, hyperlinks, and a synthesized dimension.
 * Styles, formula text, comments, sheet geometry, and images are not carried by this reader and
 * degrade to empty/undefined on their accessors. Encrypted documents and non-spreadsheet ODF
 * files fail with a typed {@link XlsxError}.
 */
export async function openOds(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<Workbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
	const zip = openZip(bytes, options); // throws `not-a-zip` when the bytes aren't a ZIP at all.

	// Encrypted ODF (an encryption-data entry in the manifest) can't be parsed — refuse it typed
	// rather than garbage-parse the ciphertext.
	if (zip.has("META-INF/manifest.xml")) {
		const manifest = decoder.decode(await zip.read("META-INF/manifest.xml"));
		if (manifest.includes("manifest:encryption-data")) {
			throw new XlsxError("unsupported", "encrypted OpenDocument files are not supported");
		}
	}

	// Container identity: when a `mimetype` entry is present it must name a spreadsheet, so an .odt
	// or .odp fed to openOds is rejected. Tolerant when absent (a few producers omit it) — the
	// presence of a parseable content.xml then decides.
	if (zip.has("mimetype")) {
		const mimetype = decoder.decode(await zip.read("mimetype")).trim();
		if (mimetype !== "" && !mimetype.startsWith(MIMETYPE_SPREADSHEET)) {
			throw new XlsxError(
				"unsupported",
				`not an OpenDocument spreadsheet (mimetype: ${mimetype})`,
			);
		}
	}

	if (!zip.has("content.xml")) {
		throw new XlsxError("missing-part", "ods is missing a required part: content.xml");
	}
	const parsed = parseOdsContent(decoder.decode(await zip.read("content.xml")));

	const infos: SheetInfo[] = [];
	const byName = new Map<string, Worksheet>();
	let index = 0;
	for (const sheet of parsed) {
		index += 1;
		const name = sheet.name !== "" ? sheet.name : `Sheet${index}`;
		const info: SheetInfo = {
			name,
			path: `content.xml#${name}`,
			visible: sheet.visible,
			state: sheet.state,
		};
		infos.push(info);
		// First definition wins if two tables somehow share a name (matches openXlsx).
		if (!byName.has(name)) {
			byName.set(
				name,
				new OdsWorksheet(
					info,
					sheet.cells,
					sheet.merges,
					sheet.hyperlinks,
					sheet.dimension,
				),
			);
		}
	}

	return new Workbook(infos, byName);
}
