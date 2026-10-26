import { inferCsvValue, MAX_CSV_CELLS, scanDelimited, sniffDelimiter } from "../csv";
import { XlsxError } from "../errors";
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
import { decodeText } from "./decode";
import { Workbook } from "./workbook";

// The `.csv` / `.tsv` reader. Delimited text is the universal export; it has no container and no
// async work, so `openCsv` is synchronous and returns the SAME public Workbook the other readers do
// (one sheet — a CSV is a single table). A hand-rolled scanner (see ../csv) handles RFC 4180 quoting
// and every line ending; type inference is conservative (numbers + booleans only, never dates), so a
// spreadsheet's worth of columns crosses into the typed model without fabricating values (F7.3).
// Fields go straight from the scanner into cells, so no grid of every field is ever built (F11.58).

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

	get protection(): undefined {
		return undefined;
	}

	get pageMargins(): undefined {
		return undefined;
	}

	get pageSetup(): undefined {
		return undefined;
	}

	get printOptions(): undefined {
		return undefined;
	}

	get headerFooter(): undefined {
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

/** The cells {@link csvCells} places, and their synthesized dimension. */
export interface CsvCells {
	readonly cells: Map<string, Cell>;
	readonly dimension: string | undefined;
}

/**
 * Scan delimited text into typed cells, keyed by A1 ref, with the dimension that spans them. Empty
 * fields are absent cells, and fields past Excel's grid are dropped. Placing more than `maxCells`
 * non-empty cells fails with `XlsxError('part-too-large')`; the default is `MAX_CSV_CELLS` (F11.58).
 *
 * @example csvCells("a,1\n,TRUE", ",", true) // { cells: A1 "a", B1 1, B2 true; dimension: "A1:B2" }
 * @example csvCells("a,b,c", ",", true, 2) // throws XlsxError part-too-large
 */
export function csvCells(
	text: string,
	delimiter: string,
	infer: boolean,
	maxCells: number = MAX_CSV_CELLS,
): CsvCells {
	const cells = new Map<string, Cell>();
	let minCol = Number.POSITIVE_INFINITY;
	let minRow = Number.POSITIVE_INFINITY;
	let maxCol = 0;
	let maxRow = 0;

	scanDelimited(text, delimiter, (field, r, c) => {
		if (r >= MAX_ROW || c >= MAX_COL) return;
		const data = infer
			? inferCsvValue(field)
			: field === ""
				? undefined
				: ({ type: "string", value: field } as const);
		if (data === undefined) return;
		if (cells.size >= maxCells) {
			throw new XlsxError(
				"part-too-large",
				`csv: the input has more than ${maxCells} non-empty cells, the cell budget`,
			);
		}
		const col = c + 1;
		const row = r + 1;
		const ref = formatRef({ col, row });
		cells.set(ref, { ref, type: data.type, value: data.value } as Cell);
		if (col < minCol) minCol = col;
		if (col > maxCol) maxCol = col;
		if (row < minRow) minRow = row;
		if (row > maxRow) maxRow = row;
	});

	const dimension =
		cells.size === 0
			? undefined
			: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`;
	return { cells, dimension };
}

/**
 * Read delimited text (`.csv` / `.tsv`) into a {@link Workbook} of one sheet — the same public
 * surface {@link openXlsx} returns. Accepts raw bytes (decoded as UTF-8) or a string. Cells carry
 * inferred number/boolean types (never dates — see {@link CsvReadOptions.inferTypes}); merges,
 * styles, formulas, comments, geometry, and images degrade on their accessors. Synchronous: CSV has
 * no container to decompress. Bytes whose text exceeds the runtime's maximum string length, and input
 * with more than 16,000,000 non-empty cells within Excel's grid, fail with
 * `XlsxError('part-too-large')` (F11.58).
 *
 * @example openCsv("name,qty\nApples,42").sheet("Sheet1").cell("B2") // { ref: "B2", type: "number", value: 42 }
 * @example openCsv(new Uint8Array(0x1fffffe8 + 8).fill(0x41)) // throws XlsxError part-too-large
 */
export function openCsv(
	source: Uint8Array | ArrayBuffer | string,
	options?: CsvReadOptions,
): Workbook {
	const text =
		typeof source === "string"
			? source
			: decodeText(source instanceof Uint8Array ? source : new Uint8Array(source));
	const delimiter =
		options?.delimiter === undefined || options.delimiter === "auto"
			? sniffDelimiter(text)
			: options.delimiter;
	const infer = options?.inferTypes ?? true;
	const sheetName = options?.sheetName ?? "Sheet1";

	const { cells, dimension } = csvCells(text, delimiter, infer);
	const info: SheetInfo = { name: sheetName, path: sheetName, visible: true, state: "visible" };
	const byName = new Map<string, Worksheet>([
		[sheetName, new CsvWorksheet(info, cells, dimension)],
	]);
	return new Workbook([info], byName);
}
