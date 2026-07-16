import { inferCsvValue, parseDelimited, sniffDelimiter } from "../csv";
import { formatRef, MAX_COL, MAX_ROW, parseRef } from "../ooxml";
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
import { Workbook } from "./workbook";

// The `.csv` / `.tsv` reader. Delimited text is the universal export; it has no container and no
// async work, so `openCsv` is synchronous and returns the SAME public Workbook the other readers do
// (one sheet — a CSV is a single table). A hand-rolled scanner (see ../csv) handles RFC 4180 quoting
// and every line ending; type inference is conservative (numbers + booleans only, never dates), so a
// spreadsheet's worth of columns crosses into the typed model without fabricating values (F7.3).

const decoder = new TextDecoder();

const NO_MERGES: readonly string[] = [];
const NO_HYPERLINKS: readonly Hyperlink[] = [];
const NO_COMMENTS: readonly Comment[] = [];
const NO_TABLES: readonly TableInfo[] = [];
const NO_DATA_VALIDATIONS: readonly DataValidation[] = [];
const NO_CONDITIONAL_FORMATTING: readonly ConditionalFormatting[] = [];
const NO_COLUMNS: readonly ColumnProps[] = [];
const NO_IMAGES: readonly SheetImage[] = [];
const NO_ROW_PROPS: ReadonlyMap<number, RowProps> = new Map();

/**
 * Options for {@link openCsv}. CSV genuinely needs parse-time knobs the container formats don't, so
 * this is an additive, documented exception to the one-options-type rule.
 */
export interface CsvReadOptions {
	/** Field delimiter. `"auto"` (default) sniffs comma / tab / semicolon from the first line. */
	readonly delimiter?: "," | "\t" | ";" | "auto";
	/** The single sheet's name. Defaults to `"Sheet1"`. */
	readonly sheetName?: string;
	/** Infer number/boolean types (default `true`); `false` reads every non-empty field as a string. */
	readonly inferTypes?: boolean;
}

// A CSV-backed worksheet: a plain data holder (cells + a synthesized dimension). Everything else the
// format can't express degrades to the shared empties — the F7.1 `OdsWorksheet` pattern.
class CsvWorksheet implements Worksheet {
	readonly name: string;
	readonly #info: SheetInfo;
	readonly #cells: Map<string, Cell>;
	readonly #dimension: string | undefined;
	#rowsCache: Row[] | undefined;

	constructor(info: SheetInfo, cells: Map<string, Cell>, dimension: string | undefined) {
		this.name = info.name;
		this.#info = info;
		this.#cells = cells;
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

	get dimension(): string | undefined {
		return this.#dimension;
	}

	// ── CSV carries only cell values; everything else degrades. ──
	get mergedCells(): readonly string[] {
		return NO_MERGES;
	}

	get hyperlinks(): readonly Hyperlink[] {
		return NO_HYPERLINKS;
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

	get autoFilter(): undefined {
		return undefined;
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
 * Read delimited text (`.csv` / `.tsv`) into a {@link Workbook} of one sheet — the same public
 * surface {@link openXlsx} returns. Accepts raw bytes (decoded as UTF-8) or a string. Cells carry
 * inferred number/boolean types (never dates — see {@link CsvReadOptions.inferTypes}); merges,
 * styles, formulas, comments, geometry, and images degrade on their accessors. Synchronous: CSV has
 * no container to decompress.
 */
export function openCsv(
	source: Uint8Array | ArrayBuffer | string,
	options?: CsvReadOptions,
): Workbook {
	const text =
		typeof source === "string"
			? source
			: decoder.decode(source instanceof Uint8Array ? source : new Uint8Array(source));
	const delimiter =
		options?.delimiter === undefined || options.delimiter === "auto"
			? sniffDelimiter(text)
			: options.delimiter;
	const infer = options?.inferTypes ?? true;
	const sheetName = options?.sheetName ?? "Sheet1";

	const grid = parseDelimited(text, delimiter);
	const cells = new Map<string, Cell>();
	let minCol = Number.POSITIVE_INFINITY;
	let minRow = Number.POSITIVE_INFINITY;
	let maxCol = 0;
	let maxRow = 0;

	const rowCount = Math.min(grid.length, MAX_ROW);
	for (let r = 0; r < rowCount; r++) {
		const fields = grid[r] as string[];
		const colCount = Math.min(fields.length, MAX_COL);
		for (let c = 0; c < colCount; c++) {
			const field = fields[c] as string;
			const data = infer
				? inferCsvValue(field)
				: field === ""
					? undefined
					: ({ type: "string", value: field } as const);
			if (data === undefined) continue;
			const col = c + 1;
			const row = r + 1;
			const ref = formatRef({ col, row });
			cells.set(ref, { ref, type: data.type, value: data.value } as Cell);
			if (col < minCol) minCol = col;
			if (col > maxCol) maxCol = col;
			if (row < minRow) minRow = row;
			if (row > maxRow) maxRow = row;
		}
	}

	const dimension =
		cells.size === 0
			? undefined
			: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`;
	const info: SheetInfo = { name: sheetName, path: sheetName, visible: true, state: "visible" };
	const byName = new Map<string, Worksheet>([
		[sheetName, new CsvWorksheet(info, cells, dimension)],
	]);
	return new Workbook([info], byName);
}
