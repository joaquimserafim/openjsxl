import { XlsxError } from "../errors";
import {
	formatRef,
	MAX_COL,
	MAX_COL_WIDTH,
	MAX_ROW,
	MAX_ROW_HEIGHT,
	parseCanonicalCell,
	parseCanonicalRange,
	rangeRunsForward,
} from "../ooxml/a1";
import {
	colorScaleCountsOk,
	dataBarCountsOk,
	iconSetCount,
	iconSetCountsOk,
	MAX_CF_FORMULAS,
} from "../ooxml/conditional-formatting";
import {
	DATA_VALIDATION_ERROR_STYLES,
	DATA_VALIDATION_OPERATORS,
	DATA_VALIDATION_TYPES,
	isCanonicalSqrefToken,
	isDataValidationErrorStyle,
	isDataValidationOperator,
	isDataValidationType,
	MAX_DV_TEXT_LEN,
	MAX_DV_TITLE_LEN,
	MAX_SQREF_RANGES,
} from "../ooxml/data-validation";
import { dateToSerial } from "../ooxml/dates";
import { MAX_EMU, MEDIA_MIME_TO_EXT } from "../ooxml/drawing";
import { MAX_FORMULA_LEN } from "../ooxml/formula";
import {
	MAX_PAGE_MARGIN,
	MAX_PAGE_SCALE,
	MAX_PAGE_SETUP_UINT,
	MAX_SPIN_COUNT,
	MIN_PAGE_SCALE,
} from "../ooxml/styles";
import { MAX_TABLE_NAME_LEN, type TableNameProblem, tableNameProblem } from "../ooxml/table";
import { encodeXstring } from "../ooxml/xstring";
import type {
	CellStyle,
	ColumnProps,
	ConditionalFormatting,
	DataValidation,
	DataValidationType,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetImage,
	TableInfo,
} from "../types";
import { commentsParts } from "./comments";
import type { MediaRegistry } from "./images";
import { checkKeys, sheetInvalid, shortened } from "./sheet-validation";
import { colorXml, type Fail, type StyleRegistry, validateColor } from "./styles";
import type { CellInput, CellValue, SheetInput, StreamSheetInput, StyledCell } from "./types";
import {
	escapeAttr,
	escapeText,
	isPlainRecord,
	isXmlSafe,
	ownValue,
	preserveAttr,
	rejectPrototypeFields,
} from "./xml";

const encoder = new TextEncoder();

// Serialize one sheet's rows into worksheet XML (`xl/worksheets/sheetN.xml`). The element order the
// schema requires here is <dimension> then <sheetData>; within <sheetData>, rows ascend by index and
// cells ascend by column — which is exactly the order we walk the input arrays.
//
// Styles (F4.2): every cell resolves an xf index through the shared StyleRegistry — 0 (the default)
// omits the `s` attribute entirely, so bare-value input emits the exact pre-F4.2 bytes. A styled
// BLANK cell ({ value: null, style }) is real: it emits a valueless `<c r s/>` and counts toward
// the dimension, which is how a border or fill lands on an empty cell.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// A finite JS number as it appears in <v>. String() gives the shortest decimal that parses back to
// the same double via the reader's Number() — so the value round-trips exactly. Non-finite values
// (NaN, ±Infinity) have no .xlsx representation and are rejected before this.
function numberToXml(n: number): string {
	return String(n);
}

// Caller properties captured for one cell occurrence. Style registration stays in row order;
// buffered table headers may retain this record until their cells are emitted.
export interface CapturedCell {
	readonly value: CellValue;
	readonly formula: unknown;
	readonly style: CellStyle | undefined;
}

// The StyledCell fields, checked against the type so a new field cannot be left out. A cell's
// prototype must not supply any of them (F11.34).
const CELL_FIELDS = Object.keys({
	value: true,
	style: true,
	formula: true,
} satisfies Record<keyof StyledCell, true>);

/**
 * Split a CellInput into its value, (optional) style, and (optional) formula. Discrimination is
 * total: null/undefined are empty, Date instances are dates, primitives are bare values — any OTHER
 * object must be a StyledCell. An object with neither a `value` nor a `formula` property is some
 * stray object (the pre-F4.2 writer rejected every object; keeping that loudness catches typos like
 * { val: 1 } or a nested array). A class instance or `Object.create` record whose prototype supplies
 * `value`, `style` or `formula` is rejected rather than written without it (F11.34). Each caller
 * property is read exactly once (TOCTOU-safe).
 *
 * @example splitInput(1, 1, null) // { value: null, formula: undefined, style: undefined }
 * @example splitInput(1, 1, 5) // { value: 5, formula: undefined, style: undefined }
 * @example splitInput(1, 1, [5] as never) // throws XlsxError invalid-input: cell A1: an array is not a cell value
 * @example splitInput(1, 1, new (class { value = 5; get style() { return {}; } })()) // throws XlsxError invalid-input: cell A1: "style" comes from a prototype; set it as an own property
 * @example splitInput(1, 1, { value: 5, typo: 1 } as never) // throws XlsxError invalid-input: cell A1: a cell object allows only "value", "style", and "formula" (got "typo")
 * @example splitInput(1, 1, {}) // throws XlsxError invalid-input: cell A1: an object cell must be { value, style? } or carry a formula
 * @example splitInput(1, 1, { value: { value: 5 } } as never) // throws XlsxError invalid-input: cell A1: a cell's value cannot be an object
 * @example splitInput(1, 1, { formula: "1+1", value: 2, style: { font: { bold: true } } }) // { value: 2, formula: "1+1", style: { font: { bold: true } } }
 */
export function splitInput(col: number, row: number, input: CellInput): CapturedCell {
	if (input === null || input === undefined) {
		return { value: input, formula: undefined, style: undefined };
	}
	if (typeof input !== "object" || input instanceof Date) {
		return { value: input, formula: undefined, style: undefined };
	}
	const ref = formatRef({ col, row });
	if (Array.isArray(input)) {
		throw new XlsxError("invalid-input", `cell ${ref}: an array is not a cell value`);
	}
	rejectPrototypeFields(input, CELL_FIELDS, `cell ${ref}`);
	// `input` is a StyledCell here — every other CellInput shape returned above — so read it directly,
	// no cast. Its properties are still caller-controlled: read each once (TOCTOU) and re-validate,
	// since a JS caller can pass anything past the types.
	for (const key of Object.keys(input)) {
		if (key !== "value" && key !== "style" && key !== "formula") {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: a cell object allows only "value", "style", and "formula" (got "${key}")`,
			);
		}
	}
	const hasFormula = Object.hasOwn(input, "formula");
	if (!hasFormula && !Object.hasOwn(input, "value")) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: an object cell must be { value, style? } or carry a formula`,
		);
	}
	const value = ownValue(input, "value");
	// The inner value must be a plain CellValue — a nested { value } or any other object (except Date)
	// has no meaning and would silently mis-serialize. This guard is RUNTIME-only: a JS caller can pass
	// a non-Date object even though the type forbids it, so keep it even though TS sees it as dead.
	if (
		value !== null &&
		value !== undefined &&
		typeof value === "object" &&
		!(value instanceof Date)
	) {
		throw new XlsxError("invalid-input", `cell ${ref}: a cell's value cannot be an object`);
	}
	return {
		value,
		formula: hasFormula ? ownValue(input, "formula") : undefined,
		style: ownValue(input, "style"),
	};
}

/**
 * Render a single cell, or `undefined` for one that produces no output (an empty value with no
 * effective style). Throws `invalid-input` for a value that cannot be represented: a non-finite
 * number, an invalid Date, or a type outside the CellInput union (JS callers can pass anything).
 * `date1904` selects the serial epoch so it matches the workbook's declared <workbookPr date1904>.
 * An empty cell with no style and no formula returns before its A1 ref is built (F11.54).
 *
 * Each example passes a fresh `createStyleRegistry()` as `styles`:
 * @example renderCell(1, 1, { value: undefined, formula: undefined, style: undefined }, false, styles) // undefined
 * @example renderCell(2, 1, { value: null, formula: undefined, style: { font: { bold: true } } }, false, styles) // '<c r="B1" s="1"/>'
 * @example renderCell(1, 1, { value: "a", formula: undefined, style: undefined }, false, styles) // '<c r="A1" t="inlineStr"><is><t>a</t></is></c>'
 * @example renderCell(1, 1, { value: true, formula: undefined, style: undefined }, false, styles) // '<c r="A1" t="b"><v>1</v></c>'
 * @example renderCell(1, 1, { value: 2.5, formula: undefined, style: undefined }, false, styles) // '<c r="A1"><v>2.5</v></c>'
 * @example renderCell(1, 1, { value: new Date(Date.UTC(2020, 0, 1)), formula: undefined, style: undefined }, false, styles) // '<c r="A1" s="1"><v>43831</v></c>'
 * @example renderCell(1, 1, { value: 2, formula: "1+1", style: undefined }, false, styles) // '<c r="A1"><f>1+1</f><v>2</v></c>'
 * @example renderCell(1, 1, { value: Number.NaN, formula: undefined, style: undefined }, false, styles) // throws XlsxError invalid-input: cell A1: NaN is not a finite number
 * @example renderCell(1, 1, { value: new Date(Number.NaN), formula: undefined, style: undefined }, false, styles) // throws XlsxError invalid-input: cell A1: invalid Date
 * @example renderCell(1, 1, { value: 1n, formula: undefined, style: undefined }, false, styles) // throws XlsxError invalid-input: cell A1: unsupported cell value type
 */
export function renderCell(
	col: number,
	row: number,
	cell: CapturedCell,
	date1904: boolean,
	styles: StyleRegistry,
): string | undefined {
	const { value, style, formula } = cell;
	if (formula === undefined && style === undefined && (value === null || value === undefined)) {
		return undefined;
	}
	const ref = formatRef({ col, row });

	// Resolve the xf index. Bare non-date values never touch the registry (zero overhead on the
	// unstyled path); a Date always does (it needs the date number format).
	let xf = 0;
	if (value instanceof Date) {
		xf = styles.xfIndexFor(style, true, ref);
	} else if (style !== undefined) {
		xf = styles.xfIndexFor(style, false, ref);
	}
	const sAttr = xf === 0 ? "" : ` s="${xf}"`;

	// A formula cell (F5.4) emits `<c s?><f>…</f><v>cached</v></c>`: the formula plus its optional
	// cached result. The result determines the cell type exactly like a bare value, except a string
	// result uses `t="str"` (a formula string) rather than an inline string.
	if (formula !== undefined) {
		return renderFormulaCell(ref, sAttr, formula, value, date1904);
	}

	if (value === null || value === undefined) {
		// A styled blank emits a valueless cell; an unstyled (or default-styled) empty is omitted.
		return xf === 0 ? undefined : `<c r="${ref}"${sAttr}/>`;
	}
	if (typeof value === "string") {
		// String content is ST_Xstring (F9.6): characters XML cannot carry (controls, lone
		// surrogates) store as `_xHHHH_`, and a literal look-alike is protected as `_x005F_…` —
		// the convention Excel/openpyxl decode on load. A clean string passes through unchanged
		// (byte-identity); nothing is rejected here anymore, so whatever string a reader returns
		// (xlsx escape, raw xlsb/csv control char) writes losslessly.
		const stored = encodeXstring(value);
		return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${preserveAttr(stored)}>${escapeText(stored)}</t></is></c>`;
	}
	if (typeof value === "boolean") {
		return `<c r="${ref}"${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
	}
	if (typeof value === "number") {
		if (!Number.isFinite(value)) {
			throw new XlsxError("invalid-input", `cell ${ref}: ${value} is not a finite number`);
		}
		return `<c r="${ref}"${sAttr}><v>${numberToXml(value)}</v></c>`;
	}
	if (value instanceof Date) {
		const serial = dateToSerial(value, date1904);
		if (!Number.isFinite(serial)) {
			throw new XlsxError("invalid-input", `cell ${ref}: invalid Date`);
		}
		return `<c r="${ref}"${sAttr}><v>${numberToXml(serial)}</v></c>`;
	}
	throw new XlsxError("invalid-input", `cell ${ref}: unsupported cell value type`);
}

/** Holes a row may show beyond its visited columns before the rest comes from its keys (F11.54). */
export const SPARSE_ROW_SLACK = 32;

/** A row's own column indices from one key listing, ascending, under the length they were listed for. */
export interface RowListing {
	readonly length: number;
	readonly indices: readonly number[];
}

/** Key listings kept per row object for one sheet, so a row object reused across rows is listed once. */
export type RowListings = WeakMap<object, RowListing>;

/**
 * Visit a row's columns in ascending order: `visit(column, own)` for each own index below `length`
 * and each `captured` column below `length` (a table-header cell read before emission, which
 * renders from its capture). `own` is the `Object.hasOwn` probe made just before the visit, so the
 * caller reads `cells[column]` only when it is true. Holes are skipped, so a sparse row costs its
 * populated columns, not its length (F11.54).
 *
 * Columns are probed in order, each just before its visit, so a dense row reads exactly as an
 * index loop would. Once a row's holes outnumber its visited columns by more than
 * `SPARSE_ROW_SLACK`, the rest of the row comes from one listing of its own keys, kept in
 * `listings` for that row object; each listed column is probed again just before its visit.
 *
 * @example forEachRowColumn([1, , 3], 3, visit) // visit(0, true), visit(2, true)
 * @example forEachRowColumn(Object.assign([], { 16383: "x" }), 16384, visit) // visit(16383, true), after 33 probes and one key listing
 * @example forEachRowColumn([, "h"], 2, visit, new Set([0, 5])) // visit(0, false), visit(1, true); column 5 is past the row
 * @example forEachRowColumn([], 0, visit) // no visits
 */
export function forEachRowColumn(
	cells: readonly unknown[],
	length: number,
	visit: (column: number, own: boolean) => void,
	captured?: ReadonlySet<number>,
	listings?: RowListings,
): void {
	let visited = 0;
	let holes = 0;
	let next = 0;
	for (; next < length; next++) {
		const own = Object.hasOwn(cells, next);
		if (own || captured?.has(next) === true) {
			visited++;
			visit(next, own);
		} else if (++holes > visited + SPARSE_ROW_SLACK) {
			break;
		}
	}
	if (next >= length) return;
	let listing = listings?.get(cells);
	if (listing === undefined || listing.length !== length) {
		const indices: number[] = [];
		for (const key of Object.getOwnPropertyNames(cells)) {
			const index = Number(key);
			// Only canonical array indices: "length", "01", "-0" and "1e3" are other own keys.
			if (Number.isInteger(index) && index >= 0 && index < length && String(index) === key) {
				indices.push(index);
			}
		}
		// An ordinary array lists its indices ascending; a Proxy's ownKeys need not.
		indices.sort((a, b) => a - b);
		listing = { length, indices };
		listings?.set(cells, listing);
	}
	const rest = listing.indices.filter((index) => index > next);
	if (captured !== undefined) {
		for (const index of captured) if (index > next && index < length) rest.push(index);
		rest.sort((a, b) => a - b);
	}
	for (let i = 0; i < rest.length; i++) {
		const column = rest[i] as number;
		if (i > 0 && column === rest[i - 1]) continue;
		const own = Object.hasOwn(cells, column);
		if (own || captured?.has(column) === true) visit(column, own);
	}
}

// The cached result of a formula, as a `t` attribute + `<v>` element (empty pair when there is no
// cached value — Excel computes it on open). A string result is `t="str"` (a formula string), NOT an
// inline string; other types match a bare value. `styled`'s xf already carries any date format.
function cachedValueXml(
	ref: string,
	value: CellValue,
	date1904: boolean,
): { readonly tAttr: string; readonly vXml: string } {
	if (value === null || value === undefined) return { tAttr: "", vXml: "" };
	if (typeof value === "string") {
		// A cached formula string's <v> is ST_Xstring too — encode, don't reject (F9.6), so a
		// cached result the reader decoded (or any API string) round-trips instead of aborting.
		const stored = encodeXstring(value);
		return { tAttr: ' t="str"', vXml: `<v>${escapeText(stored)}</v>` };
	}
	if (typeof value === "boolean") return { tAttr: ' t="b"', vXml: `<v>${value ? 1 : 0}</v>` };
	if (typeof value === "number") {
		if (!Number.isFinite(value)) {
			throw new XlsxError("invalid-input", `cell ${ref}: ${value} is not a finite number`);
		}
		return { tAttr: "", vXml: `<v>${numberToXml(value)}</v>` };
	}
	if (value instanceof Date) {
		const serial = dateToSerial(value, date1904);
		if (!Number.isFinite(serial))
			throw new XlsxError("invalid-input", `cell ${ref}: invalid Date`);
		return { tAttr: "", vXml: `<v>${numberToXml(serial)}</v>` };
	}
	throw new XlsxError("invalid-input", `cell ${ref}: unsupported cached value type`);
}

// Render a formula cell. `formula` is the caller's value read once (TOCTOU): it must be a non-empty,
// XML-safe string in stored form (no leading `=`) within Excel's length ceiling. It is LOGICAL text:
// ST_Formula is a restriction of ST_Xstring (F11.20), so the complete value goes through
// encodeXstring exactly once, after validation and immediately before XML escaping — a literal
// `_x0041_` stores as `_x005F_x0041_` (Excel would otherwise read it as `A`), CR stores as
// `_x000D_` (the canonical CR wire form E20 established), and clean text is byte-identical.
function renderFormulaCell(
	ref: string,
	sAttr: string,
	formula: unknown,
	value: CellValue,
	date1904: boolean,
): string {
	if (typeof formula !== "string") {
		throw new XlsxError("invalid-input", `cell ${ref}: formula must be a string`);
	}
	if (formula.length === 0) {
		throw new XlsxError("invalid-input", `cell ${ref}: formula must not be empty`);
	}
	if (formula.length > MAX_FORMULA_LEN) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: formula exceeds Excel's ${MAX_FORMULA_LEN}-character limit`,
		);
	}
	if (formula.startsWith("=")) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: formula must be in stored form, without a leading "="`,
		);
	}
	if (!isXmlSafe(formula)) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: formula contains a character not allowed in XML (a control character or lone surrogate)`,
		);
	}
	const cached = cachedValueXml(ref, value, date1904);
	return `<c r="${ref}"${sAttr}${cached.tAttr}><f>${escapeText(encodeXstring(formula))}</f>${cached.vXml}</c>`;
}

// ── Sheet geometry (F4.5): validation + emission ───────────────────────────────────────────────
// Same philosophy as styles: strict validation naming the sheet, `false`/empty normalize away,
// and the accepted bounds are exactly what the reader's geometry accessors can produce — so the
// bridge's geometry always writes.

// <cols> — one <col> per entry that survives normalization ({hidden: false} alone melts away).
function colsXml(sheetName: string, columns: readonly ColumnProps[] | undefined): string {
	if (columns === undefined) return "";
	if (!Array.isArray(columns)) sheetInvalid(sheetName, "columns must be an array");
	const entries: string[] = [];
	for (let i = 0; i < columns.length; i++) {
		const raw = ownValue(columns, i) as unknown;
		const what = `columns[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["min", "max", "width", "hidden"]);
		const min = ownValue(raw, "min");
		const max = ownValue(raw, "max");
		if (
			typeof min !== "number" ||
			!Number.isInteger(min) ||
			typeof max !== "number" ||
			!Number.isInteger(max) ||
			min < 1 ||
			max < min ||
			max > MAX_COL
		) {
			sheetInvalid(
				sheetName,
				`${what} needs integer 1-based min ≤ max within Excel's ${MAX_COL} columns`,
			);
		}
		let attrs = ` min="${min}" max="${max}"`;
		const width = ownValue(raw, "width");
		if (width !== undefined) {
			if (
				typeof width !== "number" ||
				!Number.isFinite(width) ||
				width <= 0 ||
				width > MAX_COL_WIDTH
			) {
				sheetInvalid(sheetName, `${what}.width must be a number in (0, ${MAX_COL_WIDTH}]`);
			}
			// customWidth marks the width as user-set — Excel ignores a bare width without it.
			attrs += ` width="${String(width)}" customWidth="1"`;
		}
		const hidden = ownValue(raw, "hidden");
		if (hidden !== undefined && typeof hidden !== "boolean") {
			sheetInvalid(sheetName, `${what}.hidden must be a boolean`);
		}
		if (hidden === true) attrs += ' hidden="1"';
		if (width !== undefined || hidden === true) entries.push(`<col${attrs}/>`);
	}
	return entries.length > 0 ? `<cols>${entries.join("")}</cols>` : "";
}

// Per-row `ht`/`hidden` attributes, keyed by 1-based row number. Rows whose properties all
// normalize away are dropped; the survivors may belong to rows with no cells at all.
function rowAttrsMap(
	sheetName: string,
	rowProperties: Readonly<Record<number, RowProps>> | undefined,
): Map<number, string> {
	const out = new Map<number, string>();
	if (rowProperties === undefined) return out;
	if (!isPlainRecord(rowProperties)) sheetInvalid(sheetName, "rowProperties must be an object");
	const record: Record<string, unknown> = rowProperties;
	for (const key of Object.keys(record)) {
		const rowNum = Number(key);
		if (!Number.isInteger(rowNum) || rowNum < 1 || rowNum > MAX_ROW) {
			sheetInvalid(
				sheetName,
				`rowProperties key "${key}" is not a row number within Excel's grid`,
			);
		}
		const raw = ownValue(record, key);
		const what = `rowProperties[${key}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["height", "hidden"]);
		let attrs = "";
		const height = ownValue(raw, "height");
		if (height !== undefined) {
			if (
				typeof height !== "number" ||
				!Number.isFinite(height) ||
				height <= 0 ||
				height > MAX_ROW_HEIGHT
			) {
				sheetInvalid(
					sheetName,
					`${what}.height must be a number in (0, ${MAX_ROW_HEIGHT}]`,
				);
			}
			// customHeight marks the height as user-set, mirroring customWidth on columns.
			attrs += ` ht="${String(height)}" customHeight="1"`;
		}
		const hidden = ownValue(raw, "hidden");
		if (hidden !== undefined && typeof hidden !== "boolean") {
			sheetInvalid(sheetName, `${what}.hidden must be a boolean`);
		}
		if (hidden === true) attrs += ' hidden="1"';
		if (attrs !== "") out.set(rowNum, attrs);
	}
	return out;
}

// <sheetViews> with a frozen <pane>. For state="frozen", xSplit/ySplit are whole column/row
// counts; topLeftCell is the first scrollable cell and activePane the quadrant the cursor lives
// in — Excel expects all three to be consistent.
function sheetViewsXml(sheetName: string, freeze: FreezePane | undefined): string {
	if (freeze === undefined) return "";
	if (!isPlainRecord(freeze)) sheetInvalid(sheetName, "freeze must be an object");
	checkKeys(sheetName, "freeze", freeze, ["rows", "cols"]);
	const validate = (value: unknown, what: string, limit: number): number => {
		if (value === undefined) return 0;
		if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value >= limit) {
			sheetInvalid(sheetName, `freeze.${what} must be an integer in [0, ${limit})`);
		}
		return value;
	};
	const rows = validate(ownValue(freeze, "rows"), "rows", MAX_ROW);
	const cols = validate(ownValue(freeze, "cols"), "cols", MAX_COL);
	if (rows === 0 && cols === 0) return ""; // freezing nothing is no freeze
	const splits = (cols > 0 ? ` xSplit="${cols}"` : "") + (rows > 0 ? ` ySplit="${rows}"` : "");
	const topLeft = formatRef({ col: cols + 1, row: rows + 1 });
	const activePane = rows > 0 && cols > 0 ? "bottomRight" : rows > 0 ? "bottomLeft" : "topRight";
	return (
		'<sheetViews><sheetView workbookViewId="0">' +
		`<pane${splits} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>` +
		"</sheetView></sheetViews>"
	);
}

// ── Structural metadata (F4.6): merges + hyperlinks — validation + emission ────────────────────
// Both blocks live AFTER </sheetData> (schema order: … sheetData, mergeCells, …, hyperlinks, …).
// Like geometry, unused blocks are empty strings, so a metadata-free sheet keeps its exact bytes.

interface MergeRect {
	readonly ref: string;
	readonly c1: number;
	readonly r1: number;
	readonly c2: number;
	readonly r2: number;
}

// <autoFilter> (F10.2) — the sheet-level filter range. Emitted in the CT_Worksheet slot after
// </sheetData> (and a future <sheetProtection>) and BEFORE <mergeCells>. Returns both the element XML
// and the validated canonical ref: the workbook writer turns that ref into the paired hidden
// `_xlnm._FilterDatabase` defined name. The ref is read ONCE (single-read TOCTOU) and validated as a
// canonical, in-grid A1 range — the same rule the reader uses to keep/drop a filter, so a read filter
// re-writes cleanly. Absent → empty string + undefined ref, so an unfiltered sheet stays byte-identical.
function autoFilterXml(
	sheetName: string,
	autoFilter: SheetInput["autoFilter"],
): { xml: string; ref: string | undefined } {
	if (autoFilter === undefined) return { xml: "", ref: undefined };
	if (!isPlainRecord(autoFilter)) sheetInvalid(sheetName, "autoFilter must be an object");
	for (const key of Object.keys(autoFilter)) {
		if (key !== "ref") sheetInvalid(sheetName, `autoFilter has unknown property "${key}"`);
	}
	const ref = ownValue(autoFilter, "ref");
	if (typeof ref !== "string") sheetInvalid(sheetName, "autoFilter.ref must be a string");
	const range = parseCanonicalRange(ref);
	if (range === undefined) {
		sheetInvalid(
			sheetName,
			`autoFilter.ref "${shortened(ref)}" is not a canonical A1 range like "A1:C10" within Excel's grid`,
		);
	}
	// Reject a backwards range (bottom-right before top-left), the SAME bound the reader drops on
	// (rangeRunsForward) — so a filter the reader surfaces always re-writes, and a filter it drops is
	// one the writer would refuse. `parseCanonicalRange` returns corners as written, order checked here.
	if (!rangeRunsForward(range)) {
		sheetInvalid(
			sheetName,
			`autoFilter.ref "${shortened(ref)}" must run top-left to bottom-right`,
		);
	}
	// `ref` is canonical (only A–Z, digits, ':') so it carries no XML-special characters — emitted verbatim.
	return { xml: `<autoFilter ref="${ref}"/>`, ref };
}

// <sheetProtection> (F10.3) — the sheet-level lock. Emitted in the CT_Worksheet slot after </sheetData>
// and BEFORE <autoFilter>. Every present boolean flag and any password material is carried; attribute
// order is fixed for determinism (XML attribute order is not schema-significant). Password hashes are
// emitted verbatim (escaped, XML-safety enforced) — never computed. An absent/empty value emits nothing,
// so an unprotected sheet stays byte-identical.
const SHEET_PROTECTION_BOOL_ATTRS = [
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
const SHEET_PROTECTION_STR_ATTRS = ["password", "algorithmName", "hashValue", "saltValue"] as const;

function sheetProtectionXml(sheetName: string, protection: SheetInput["protection"]): string {
	if (protection === undefined) return "";
	if (!isPlainRecord(protection)) sheetInvalid(sheetName, "protection must be an object");
	const allowed = new Set<string>([
		...SHEET_PROTECTION_BOOL_ATTRS,
		...SHEET_PROTECTION_STR_ATTRS,
		"spinCount",
	]);
	for (const key of Object.keys(protection)) {
		if (!allowed.has(key)) sheetInvalid(sheetName, `protection has unknown property "${key}"`);
	}
	let attrs = "";
	for (const flag of SHEET_PROTECTION_BOOL_ATTRS) {
		const v = ownValue(protection, flag);
		if (v === undefined) continue;
		if (typeof v !== "boolean") sheetInvalid(sheetName, `protection.${flag} must be a boolean`);
		attrs += ` ${flag}="${v ? 1 : 0}"`;
	}
	for (const attr of SHEET_PROTECTION_STR_ATTRS) {
		const v = ownValue(protection, attr);
		if (v === undefined) continue;
		if (typeof v !== "string") sheetInvalid(sheetName, `protection.${attr} must be a string`);
		if (!isXmlSafe(v)) {
			sheetInvalid(sheetName, `protection.${attr} contains a character not allowed in XML`);
		}
		attrs += ` ${attr}="${escapeAttr(v)}"`;
	}
	const spin = ownValue(protection, "spinCount");
	if (spin !== undefined) {
		// Bounded to xsd:unsignedInt (MAX_SPIN_COUNT): above it, `${spin}` would emit `1e+21` (invalid
		// integer XML). Shared bound with the reader's drop.
		if (
			typeof spin !== "number" ||
			!Number.isInteger(spin) ||
			spin < 0 ||
			spin > MAX_SPIN_COUNT
		) {
			sheetInvalid(
				sheetName,
				`protection.spinCount must be an integer in 0..${MAX_SPIN_COUNT}`,
			);
		}
		attrs += ` spinCount="${spin}"`;
	}
	return attrs === "" ? "" : `<sheetProtection${attrs}/>`;
}

// ── Print setup (F10.4): printOptions, pageMargins, pageSetup, headerFooter ────────────────────────
// Emitted in the CT_Worksheet slot BETWEEN <hyperlinks> and <drawing>, in schema order printOptions →
// pageMargins → pageSetup → headerFooter. Each stores ONLY explicitly-present values (attribute order
// fixed for determinism), so a sheet with no print setup stays byte-identical. Each value is read ONCE
// (single-read TOCTOU). Integer attrs are bounded to uint32 so `${v}` can't emit exponential notation
// (the spinCount trap); page margins are xsd:double, where exponential IS valid, so only finiteness and
// the shared [0, MAX_PAGE_MARGIN] range are enforced.

const PRINT_OPTIONS_FLAGS = [
	"horizontalCentered",
	"verticalCentered",
	"headings",
	"gridLines",
] as const;

function printOptionsXml(sheetName: string, po: SheetInput["printOptions"]): string {
	if (po === undefined) return "";
	if (!isPlainRecord(po)) sheetInvalid(sheetName, "printOptions must be an object");
	for (const key of Object.keys(po)) {
		if (!(PRINT_OPTIONS_FLAGS as readonly string[]).includes(key)) {
			sheetInvalid(sheetName, `printOptions has unknown property "${key}"`);
		}
	}
	let attrs = "";
	for (const flag of PRINT_OPTIONS_FLAGS) {
		const v = ownValue(po, flag);
		if (v === undefined) continue;
		if (typeof v !== "boolean")
			sheetInvalid(sheetName, `printOptions.${flag} must be a boolean`);
		attrs += ` ${flag}="${v ? 1 : 0}"`;
	}
	return attrs === "" ? "" : `<printOptions${attrs}/>`;
}

const PAGE_MARGIN_KEYS = ["left", "right", "top", "bottom", "header", "footer"] as const;

function pageMarginsXml(sheetName: string, m: SheetInput["pageMargins"]): string {
	if (m === undefined) return "";
	if (!isPlainRecord(m)) sheetInvalid(sheetName, "pageMargins must be an object");
	for (const key of Object.keys(m)) {
		if (!(PAGE_MARGIN_KEYS as readonly string[]).includes(key)) {
			sheetInvalid(sheetName, `pageMargins has unknown property "${key}"`);
		}
	}
	// All six are required by the schema.
	let attrs = "";
	for (const key of PAGE_MARGIN_KEYS) {
		const v = ownValue(m, key);
		if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > MAX_PAGE_MARGIN) {
			sheetInvalid(
				sheetName,
				`pageMargins.${key} must be a finite number in 0..${MAX_PAGE_MARGIN}`,
			);
		}
		attrs += ` ${key}="${v}"`;
	}
	return `<pageMargins${attrs}/>`;
}

const PAGE_SETUP_KEYS = [
	"paperSize",
	"orientation",
	"scale",
	"fitToWidth",
	"fitToHeight",
	"firstPageNumber",
	"useFirstPageNumber",
	"blackAndWhite",
	"draft",
	"cellComments",
	"pageOrder",
] as const;

function pageSetupXml(sheetName: string, ps: SheetInput["pageSetup"]): string {
	if (ps === undefined) return "";
	if (!isPlainRecord(ps)) sheetInvalid(sheetName, "pageSetup must be an object");
	for (const key of Object.keys(ps)) {
		if (!(PAGE_SETUP_KEYS as readonly string[]).includes(key)) {
			sheetInvalid(sheetName, `pageSetup has unknown property "${key}"`);
		}
	}
	const uint = (name: string, v: unknown): number => {
		if (typeof v !== "number" || !Number.isInteger(v) || v < 0 || v > MAX_PAGE_SETUP_UINT) {
			sheetInvalid(
				sheetName,
				`pageSetup.${name} must be an integer in 0..${MAX_PAGE_SETUP_UINT}`,
			);
		}
		return v;
	};
	const oneOf = (name: string, v: unknown, allowed: readonly string[]): string => {
		if (typeof v !== "string" || !allowed.includes(v)) {
			sheetInvalid(sheetName, `pageSetup.${name} must be one of: ${allowed.join(", ")}`);
		}
		return v;
	};
	const bool = (name: string, v: unknown): number => {
		if (typeof v !== "boolean") sheetInvalid(sheetName, `pageSetup.${name} must be a boolean`);
		return v ? 1 : 0;
	};
	// Attribute emission in CT_PageSetup schema order (single-read each).
	let attrs = "";
	const paperSize = ownValue(ps, "paperSize");
	if (paperSize !== undefined) attrs += ` paperSize="${uint("paperSize", paperSize)}"`;
	const scale = ownValue(ps, "scale");
	if (scale !== undefined) {
		if (
			typeof scale !== "number" ||
			!Number.isInteger(scale) ||
			scale < MIN_PAGE_SCALE ||
			scale > MAX_PAGE_SCALE
		) {
			sheetInvalid(
				sheetName,
				`pageSetup.scale must be an integer in ${MIN_PAGE_SCALE}..${MAX_PAGE_SCALE}`,
			);
		}
		attrs += ` scale="${scale}"`;
	}
	const firstPageNumber = ownValue(ps, "firstPageNumber");
	if (firstPageNumber !== undefined) {
		attrs += ` firstPageNumber="${uint("firstPageNumber", firstPageNumber)}"`;
	}
	const fitToWidth = ownValue(ps, "fitToWidth");
	if (fitToWidth !== undefined) attrs += ` fitToWidth="${uint("fitToWidth", fitToWidth)}"`;
	const fitToHeight = ownValue(ps, "fitToHeight");
	if (fitToHeight !== undefined) attrs += ` fitToHeight="${uint("fitToHeight", fitToHeight)}"`;
	const pageOrder = ownValue(ps, "pageOrder");
	if (pageOrder !== undefined) {
		attrs += ` pageOrder="${oneOf("pageOrder", pageOrder, ["downThenOver", "overThenDown"])}"`;
	}
	const orientation = ownValue(ps, "orientation");
	if (orientation !== undefined) {
		attrs += ` orientation="${oneOf("orientation", orientation, ["default", "portrait", "landscape"])}"`;
	}
	const blackAndWhite = ownValue(ps, "blackAndWhite");
	if (blackAndWhite !== undefined)
		attrs += ` blackAndWhite="${bool("blackAndWhite", blackAndWhite)}"`;
	const draft = ownValue(ps, "draft");
	if (draft !== undefined) attrs += ` draft="${bool("draft", draft)}"`;
	const cellComments = ownValue(ps, "cellComments");
	if (cellComments !== undefined) {
		attrs += ` cellComments="${oneOf("cellComments", cellComments, ["none", "asDisplayed", "atEnd"])}"`;
	}
	const useFirstPageNumber = ownValue(ps, "useFirstPageNumber");
	if (useFirstPageNumber !== undefined) {
		attrs += ` useFirstPageNumber="${bool("useFirstPageNumber", useFirstPageNumber)}"`;
	}
	return attrs === "" ? "" : `<pageSetup${attrs}/>`;
}

const HF_FLAGS = [
	"differentOddEven",
	"differentFirst",
	"scaleWithDoc",
	"alignWithMargins",
] as const;
const HF_CHILDREN = [
	"oddHeader",
	"oddFooter",
	"evenHeader",
	"evenFooter",
	"firstHeader",
	"firstFooter",
] as const;

function headerFooterXml(sheetName: string, hf: SheetInput["headerFooter"]): string {
	if (hf === undefined) return "";
	if (!isPlainRecord(hf)) sheetInvalid(sheetName, "headerFooter must be an object");
	const allowed = new Set<string>([...HF_FLAGS, ...HF_CHILDREN]);
	for (const key of Object.keys(hf)) {
		if (!allowed.has(key))
			sheetInvalid(sheetName, `headerFooter has unknown property "${key}"`);
	}
	let attrs = "";
	for (const flag of HF_FLAGS) {
		const v = ownValue(hf, flag);
		if (v === undefined) continue;
		if (typeof v !== "boolean")
			sheetInvalid(sheetName, `headerFooter.${flag} must be a boolean`);
		attrs += ` ${flag}="${v ? 1 : 0}"`;
	}
	let children = "";
	for (const key of HF_CHILDREN) {
		const v = ownValue(hf, key);
		if (v === undefined) continue;
		if (typeof v !== "string") sheetInvalid(sheetName, `headerFooter.${key} must be a string`);
		// The string carries Excel's &-codes as literal text; ST_Xstring-encode (like comment text, F9.6)
		// then escape — so a control char or `_xHHHH_`-shaped run survives Excel's decode.
		children += `<${key}>${escapeText(encodeXstring(v))}</${key}>`;
	}
	if (attrs === "" && children === "") return "";
	return children === ""
		? `<headerFooter${attrs}/>`
		: `<headerFooter${attrs}>${children}</headerFooter>`;
}

// The four print-setup elements in schema order (F10.4). Empty string when the sheet declares none.
// Takes just the four fields so both the buffered and streaming sheet inputs qualify.
type PrintSetupInput = Pick<
	SheetInput,
	"printOptions" | "pageMargins" | "pageSetup" | "headerFooter"
>;
function printSetupXml(sheetName: string, sheet: PrintSetupInput): string {
	return (
		printOptionsXml(sheetName, ownValue(sheet, "printOptions")) +
		pageMarginsXml(sheetName, ownValue(sheet, "pageMargins")) +
		pageSetupXml(sheetName, ownValue(sheet, "pageSetup")) +
		headerFooterXml(sheetName, ownValue(sheet, "headerFooter"))
	);
}

// <mergeCells> — Excel repair-prompts on malformed, single-cell, and overlapping merges, so all
// three are rejected. Overlap detection is a sweep over ranges sorted by first row: an active
// range is pruned once its last row passes, so every surviving active shares rows with the
// current one and a column intersection means overlap. The corollary keeps this near-linear even
// on adversarial bridge input: actives that DON'T overlap are column-disjoint, so the active list
// can never exceed MAX_COL entries — no O(n²) blow-up from a crafted file with a million merges.
function mergeCellsXml(sheetName: string, merges: readonly string[] | undefined): string {
	if (merges === undefined) return "";
	if (!Array.isArray(merges)) sheetInvalid(sheetName, "merges must be an array");
	const rects: MergeRect[] = [];
	for (let i = 0; i < merges.length; i++) {
		const ref = ownValue(merges, i) as unknown;
		const what = `merges[${i}]`;
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what} must be a string`);
		const colon = ref.indexOf(":");
		const from = colon === -1 ? undefined : parseCanonicalCell(ref.slice(0, colon));
		const to = colon === -1 ? undefined : parseCanonicalCell(ref.slice(colon + 1));
		if (from === undefined || to === undefined) {
			sheetInvalid(
				sheetName,
				`${what} "${shortened(ref)}" is not a canonical A1 range like "A1:B2" within Excel's grid`,
			);
		}
		if (to.col < from.col || to.row < from.row) {
			sheetInvalid(
				sheetName,
				`${what} "${shortened(ref)}" must run top-left to bottom-right`,
			);
		}
		if (to.col === from.col && to.row === from.row) {
			sheetInvalid(sheetName, `${what} "${shortened(ref)}" merges a single cell`);
		}
		rects.push({ ref, c1: from.col, r1: from.row, c2: to.col, r2: to.row });
	}
	if (rects.length === 0) return "";
	const sorted = [...rects].sort((a, b) => a.r1 - b.r1);
	const active: MergeRect[] = [];
	for (const rect of sorted) {
		let kept = 0;
		for (const a of active) {
			if (a.r2 < rect.r1) continue; // fully above the current range — can never overlap again
			active[kept++] = a;
			if (a.c1 <= rect.c2 && rect.c1 <= a.c2) {
				sheetInvalid(
					sheetName,
					`merges "${a.ref}" and "${rect.ref}" overlap — Excel repairs overlapping merges`,
				);
			}
		}
		active.length = kept;
		active.push(rect);
	}
	// Emission preserves input order (document order round-trips through the reader verbatim).
	return `<mergeCells count="${rects.length}">${rects
		.map((r) => `<mergeCell ref="${r.ref}"/>`)
		.join("")}</mergeCells>`;
}

// <hyperlinks> — each link covers a cell or range and needs somewhere to go: an external target
// (URL / mailto: / file:, written VERBATIM into the sheet's rels part with TargetMode="External")
// and/or a location. Without a target, the location is an in-workbook reference; with a target,
// it is a fragment within that external destination. Empty-string destinations normalize away
// (the reader never yields an empty location); tooltip/display are carried verbatim, empty or not,
// mirroring the reader exactly. Returns external targets in r:id order — non-empty means the sheet
// needs a relationships part.
function hyperlinksXml(
	sheetName: string,
	hyperlinks: readonly Hyperlink[] | undefined,
): { readonly xml: string; readonly targets: readonly string[] } {
	if (hyperlinks === undefined) return { xml: "", targets: [] };
	if (!Array.isArray(hyperlinks)) sheetInvalid(sheetName, "hyperlinks must be an array");
	const targets: string[] = [];
	const entries: string[] = [];
	for (let i = 0; i < hyperlinks.length; i++) {
		const raw = ownValue(hyperlinks, i) as unknown;
		const what = `hyperlinks[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["ref", "target", "location", "tooltip", "display"]);
		const ref = ownValue(raw, "ref");
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const colon = ref.indexOf(":");
		const from = parseCanonicalCell(colon === -1 ? ref : ref.slice(0, colon));
		const to = colon === -1 ? from : parseCanonicalCell(ref.slice(colon + 1));
		if (from === undefined || to === undefined) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" is not a canonical A1 cell or range within Excel's grid`,
			);
		}
		if (to.col < from.col || to.row < from.row) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" must run top-left to bottom-right`,
			);
		}
		// One optional string property, read exactly once (TOCTOU) and gated for XML safety.
		const str = (key: string, value: unknown): string | undefined => {
			if (value === undefined) return undefined;
			if (typeof value !== "string")
				sheetInvalid(sheetName, `${what}.${key} must be a string`);
			if (!isXmlSafe(value)) {
				sheetInvalid(
					sheetName,
					`${what}.${key} contains a character not allowed in XML (a control character or lone surrogate)`,
				);
			}
			return value;
		};
		const rawTarget = str("target", ownValue(raw, "target"));
		const rawLocation = str("location", ownValue(raw, "location"));
		// An empty destination is no destination — the reader drops location="" the same way.
		const target = rawTarget === "" ? undefined : rawTarget;
		const location = rawLocation === "" ? undefined : rawLocation;
		if (target === undefined && location === undefined) {
			sheetInvalid(sheetName, `${what} needs an external target and/or a location`);
		}
		const tooltip = str("tooltip", ownValue(raw, "tooltip"));
		const display = str("display", ownValue(raw, "display"));
		let attrs = ` ref="${ref}"`;
		if (target !== undefined) {
			targets.push(target);
			attrs += ` r:id="rId${targets.length}"`;
		}
		// `location` is logical ST_Xstring text. Encode the complete value once before XML attribute
		// escaping: `'_x0041_'!A1` becomes `'_x005F_x0041_'!A1`; CR becomes `_x000D_`, while
		// TAB/LF become numeric character references in `escapeAttr`. The previous bytes stay unchanged
		// only when the logical value contains neither escape/collision-shaped text nor CR.
		if (location !== undefined) attrs += ` location="${escapeAttr(encodeXstring(location))}"`;
		if (tooltip !== undefined) attrs += ` tooltip="${escapeAttr(tooltip)}"`;
		if (display !== undefined) attrs += ` display="${escapeAttr(display)}"`;
		entries.push(`<hyperlink${attrs}/>`);
	}
	if (entries.length === 0) return { xml: "", targets: [] };
	return { xml: `<hyperlinks>${entries.join("")}</hyperlinks>`, targets };
}

// ── Data validation (F9.2) ─────────────────────────────────────────────────────────────────────
// `<dataValidations>` slots between <mergeCells> and <hyperlinks> (CT_Worksheet order, decision 1).
// Like every structural block it is an empty string when unused, so a validation-free sheet keeps its
// exact pre-F9.2 bytes. Strict validation mirrors hyperlinks: every caller property is read once
// (TOCTOU), unknown keys and out-of-bounds values are rejected typed, bounds come from the SAME
// constants the tolerant reader clamps with, and every emitted string passes escapeAttr/escapeText +
// isXmlSafe. sqref stays symbolic — validated per range, capped in count, never expanded to cells.

// One `@sqref` token: a canonical A1 cell ("C1") or top-left:bottom-right range ("A1:A10") within
// Excel's grid. Kept symbolic; never expanded. Uses the SAME predicate the reader drops non-canonical
// tokens with (isCanonicalSqrefToken) — so a token the reader can return is always one the writer
// accepts (shared bounds).
function validateSqrefToken(sheetName: string, what: string, token: string): void {
	if (!isCanonicalSqrefToken(token)) {
		sheetInvalid(
			sheetName,
			`${what} "${shortened(token)}" is not a canonical A1 cell or top-left→bottom-right range within Excel's grid`,
		);
	}
}

function dataValidationsXml(
	sheetName: string,
	dataValidations: readonly DataValidation[] | undefined,
): string {
	if (dataValidations === undefined) return "";
	if (!Array.isArray(dataValidations))
		sheetInvalid(sheetName, "dataValidations must be an array");
	if (dataValidations.length === 0) return "";

	const entries: string[] = [];
	for (let i = 0; i < dataValidations.length; i++) {
		const raw = ownValue(dataValidations, i) as unknown;
		const what = `dataValidations[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, [
			"sqref",
			"type",
			"operator",
			"formula1",
			"formula2",
			"allowBlank",
			"showDropDown",
			"showInputMessage",
			"showErrorMessage",
			"errorStyle",
			"promptTitle",
			"prompt",
			"errorTitle",
			"error",
		]);

		// sqref — required, non-empty, canonical, capped in RANGE COUNT (decision 5).
		const sqref = ownValue(raw, "sqref");
		if (!Array.isArray(sqref)) {
			sheetInvalid(sheetName, `${what}.sqref must be an array of A1 ranges`);
		}
		if (sqref.length === 0)
			sheetInvalid(sheetName, `${what}.sqref must cover at least one range`);
		if (sqref.length > MAX_SQREF_RANGES) {
			sheetInvalid(sheetName, `${what}.sqref has more than ${MAX_SQREF_RANGES} ranges`);
		}
		const tokens: string[] = [];
		for (let s = 0; s < sqref.length; s++) {
			const token = ownValue(sqref, s) as unknown;
			if (typeof token !== "string") {
				sheetInvalid(sheetName, `${what}.sqref[${s}] must be a string`);
			}
			validateSqrefToken(sheetName, `${what}.sqref[${s}]`, token);
			tokens.push(token);
		}

		// One optional boolean property, read exactly once (TOCTOU).
		const boolValue = (key: string, value: unknown): boolean | undefined => {
			if (value === undefined) return undefined;
			if (typeof value !== "boolean")
				sheetInvalid(sheetName, `${what}.${key} must be a boolean`);
			return value;
		};
		// One optional prompt/error string: length-bounded (code points, matching the reader's clamp)
		// and XML-safe.
		const textValue = (key: string, value: unknown, max: number): string | undefined => {
			if (value === undefined) return undefined;
			if (typeof value !== "string")
				sheetInvalid(sheetName, `${what}.${key} must be a string`);
			if ([...value].length > max) {
				sheetInvalid(sheetName, `${what}.${key} exceeds ${max} characters`);
			}
			if (!isXmlSafe(value)) {
				sheetInvalid(
					sheetName,
					`${what}.${key} contains a character not allowed in XML (a control character or lone surrogate)`,
				);
			}
			return value;
		};
		// A formula operand: stored form (a leading "=" is stripped, decision 6), XML-safe. Empty is
		// no operand — the reader treats an empty <formula1> the same way.
		const formulaValue = (key: string, value: unknown): string | undefined => {
			if (value === undefined) return undefined;
			if (typeof value !== "string")
				sheetInvalid(sheetName, `${what}.${key} must be a string`);
			const stored = value.startsWith("=") ? value.slice(1) : value;
			if (stored === "") return undefined;
			if (!isXmlSafe(stored)) {
				sheetInvalid(
					sheetName,
					`${what}.${key} contains a character not allowed in XML (a control character or lone surrogate)`,
				);
			}
			return stored;
		};

		// type — default "none".
		const rawType = ownValue(raw, "type");
		let type: DataValidationType = "none";
		if (rawType !== undefined) {
			if (!isDataValidationType(rawType)) {
				sheetInvalid(
					sheetName,
					`${what}.type must be one of ${DATA_VALIDATION_TYPES.join(", ")}`,
				);
			}
			type = rawType;
		}

		const formula1 = formulaValue("formula1", ownValue(raw, "formula1"));
		const formula2 = formulaValue("formula2", ownValue(raw, "formula2"));
		// An inline list literal (`"a,b,c"`) is capped at Excel's source ceiling (decision 5). A
		// range/reference source is unbounded here.
		if (
			type === "list" &&
			formula1 !== undefined &&
			formula1.startsWith('"') &&
			[...formula1].length > MAX_DV_TEXT_LEN
		) {
			sheetInvalid(
				sheetName,
				`${what}.formula1 inline list exceeds ${MAX_DV_TEXT_LEN} characters`,
			);
		}

		let attrs = ` type="${type}"`;
		const operator = ownValue(raw, "operator");
		if (operator !== undefined) {
			if (!isDataValidationOperator(operator)) {
				sheetInvalid(
					sheetName,
					`${what}.operator must be one of ${DATA_VALIDATION_OPERATORS.join(", ")}`,
				);
			}
			attrs += ` operator="${operator}"`;
		}
		const allowBlank = boolValue("allowBlank", ownValue(raw, "allowBlank"));
		if (allowBlank !== undefined) attrs += ` allowBlank="${allowBlank ? 1 : 0}"`;
		const showInputMessage = boolValue("showInputMessage", ownValue(raw, "showInputMessage"));
		if (showInputMessage !== undefined)
			attrs += ` showInputMessage="${showInputMessage ? 1 : 0}"`;
		const showErrorMessage = boolValue("showErrorMessage", ownValue(raw, "showErrorMessage"));
		if (showErrorMessage !== undefined)
			attrs += ` showErrorMessage="${showErrorMessage ? 1 : 0}"`;
		// showDropDown is INVERTED in the file — intuitive `true` (arrow shown) writes as "0".
		const showDropDown = boolValue("showDropDown", ownValue(raw, "showDropDown"));
		if (showDropDown !== undefined) attrs += ` showDropDown="${showDropDown ? 0 : 1}"`;
		const errorStyle = ownValue(raw, "errorStyle");
		if (errorStyle !== undefined) {
			if (!isDataValidationErrorStyle(errorStyle)) {
				sheetInvalid(
					sheetName,
					`${what}.errorStyle must be one of ${DATA_VALIDATION_ERROR_STYLES.join(", ")}`,
				);
			}
			attrs += ` errorStyle="${errorStyle}"`;
		}
		const promptTitle = textValue(
			"promptTitle",
			ownValue(raw, "promptTitle"),
			MAX_DV_TITLE_LEN,
		);
		if (promptTitle !== undefined) attrs += ` promptTitle="${escapeAttr(promptTitle)}"`;
		const prompt = textValue("prompt", ownValue(raw, "prompt"), MAX_DV_TEXT_LEN);
		if (prompt !== undefined) attrs += ` prompt="${escapeAttr(prompt)}"`;
		const errorTitle = textValue("errorTitle", ownValue(raw, "errorTitle"), MAX_DV_TITLE_LEN);
		if (errorTitle !== undefined) attrs += ` errorTitle="${escapeAttr(errorTitle)}"`;
		const error = textValue("error", ownValue(raw, "error"), MAX_DV_TEXT_LEN);
		if (error !== undefined) attrs += ` error="${escapeAttr(error)}"`;
		// sqref is the LAST attribute (matching Excel/openpyxl); tokens are canonical A1 — no escaping.
		attrs += ` sqref="${tokens.join(" ")}"`;

		// Operands are logical ST_Formula (⊂ ST_Xstring) text, validated above as such; the complete
		// value is encoded once right before escaping (F11.20) — clean operands stay byte-identical.
		let children = "";
		if (formula1 !== undefined)
			children += `<formula1>${escapeText(encodeXstring(formula1))}</formula1>`;
		if (formula2 !== undefined)
			children += `<formula2>${escapeText(encodeXstring(formula2))}</formula2>`;
		entries.push(
			children === ""
				? `<dataValidation${attrs}/>`
				: `<dataValidation${attrs}>${children}</dataValidation>`,
		);
	}
	return `<dataValidations count="${entries.length}">${entries.join("")}</dataValidations>`;
}

// ── Conditional formatting (F9.3) ──────────────────────────────────────────────────────────────
// `<conditionalFormatting>` blocks slot between <mergeCells> and <dataValidations> (CT_Worksheet
// order, decision 1). A highlight rule's look is a DxfStyle interned into styles.xml's `<dxfs>` via the
// shared StyleRegistry (the numeric dxfId is assigned here, never public). Priorities are renumbered
// densely 1..n by ascending caller priority with document order as the tie-break (decision 6) — NEVER
// by position, which would silently swap which overlapping rule wins. Empty ⇒ "" (byte-identity).

const CF_HIGHLIGHT_TYPES: ReadonlySet<string> = new Set([
	"cellIs",
	"expression",
	"top10",
	"aboveAverage",
	"uniqueValues",
	"duplicateValues",
	"containsText",
	"notContainsText",
	"beginsWith",
	"endsWith",
	"containsBlanks",
	"notContainsBlanks",
	"containsErrors",
	"notContainsErrors",
	"timePeriod",
]);

// A validated rule ready to emit — everything except its (renumbered) priority.
interface BuiltCfRule {
	readonly block: number;
	readonly doc: number;
	readonly priority: number; // the caller's priority, for the sort
	readonly open: string; // `<cfRule type=... dxfId?=...` (priority inserted at emit)
	readonly rest: string; // remaining attributes
	readonly children: string;
}

/**
 * The sheet's `<conditionalFormatting>` elements: each block's validated rules, in document order,
 * with priorities renumbered densely 1..n across the sheet by (caller priority, document order).
 * Each block's rules are gathered in one pass after they are built, so emission is linear in rules
 * plus blocks (F11.55), and it walks the blocks that were validated.
 *
 * @example conditionalFormattingXml("S", undefined, styles) // ""
 * @example conditionalFormattingXml("S", [], styles) // ""
 * @example conditionalFormattingXml("S", [{ sqref: ["A1"], rules: [{ type: "cellIs", priority: 1, operator: "greaterThan", formulas: ["1"] }] }], styles) // '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>'
 * @example conditionalFormattingXml("S", {}, styles) // throws XlsxError invalid-input: conditionalFormatting must be an array
 * @example conditionalFormattingXml("S", [{ sqref: ["A1"], rules: [] }], styles) // throws XlsxError invalid-input: conditionalFormatting[0].rules must have at least one rule
 */
export function conditionalFormattingXml(
	sheetName: string,
	cfs: readonly ConditionalFormatting[] | undefined,
	styles: StyleRegistry,
): string {
	if (cfs === undefined) return "";
	if (!Array.isArray(cfs)) sheetInvalid(sheetName, "conditionalFormatting must be an array");
	if (cfs.length === 0) return "";

	const built: BuiltCfRule[] = [];
	const blockSqref: string[][] = [];
	let doc = 0;

	for (let b = 0; b < cfs.length; b++) {
		const rawBlock = ownValue(cfs, b) as unknown;
		const wb = `conditionalFormatting[${b}]`;
		if (!isPlainRecord(rawBlock)) sheetInvalid(sheetName, `${wb} must be an object`);
		checkKeys(sheetName, wb, rawBlock, ["sqref", "rules"]);

		const sqref = ownValue(rawBlock, "sqref");
		if (!Array.isArray(sqref))
			sheetInvalid(sheetName, `${wb}.sqref must be an array of A1 ranges`);
		if (sqref.length === 0)
			sheetInvalid(sheetName, `${wb}.sqref must cover at least one range`);
		if (sqref.length > MAX_SQREF_RANGES) {
			sheetInvalid(sheetName, `${wb}.sqref has more than ${MAX_SQREF_RANGES} ranges`);
		}
		const tokens: string[] = [];
		for (let s = 0; s < sqref.length; s++) {
			const tok = ownValue(sqref, s) as unknown;
			if (typeof tok !== "string")
				sheetInvalid(sheetName, `${wb}.sqref[${s}] must be a string`);
			if (!isCanonicalSqrefToken(tok)) {
				sheetInvalid(
					sheetName,
					`${wb}.sqref[${s}] "${shortened(tok)}" is not a canonical A1 cell or range within Excel's grid`,
				);
			}
			tokens.push(tok);
		}
		blockSqref[b] = tokens;

		const rules = ownValue(rawBlock, "rules");
		if (!Array.isArray(rules)) sheetInvalid(sheetName, `${wb}.rules must be an array`);
		if (rules.length === 0) sheetInvalid(sheetName, `${wb}.rules must have at least one rule`);
		for (let r = 0; r < rules.length; r++) {
			built.push(
				buildCfRule(sheetName, `${wb}.rules[${r}]`, ownValue(rules, r), styles, b, doc++),
			);
		}
	}

	// Each validated block's rules, in document order, gathered in one pass (F11.55).
	const blockRules: BuiltCfRule[][] = blockSqref.map(() => []);
	for (const rule of built) blockRules[rule.block]?.push(rule);

	// Renumber priorities densely 1..n by (caller priority asc, document order asc) — decision 6.
	const order = [...built].sort((a, z) => a.priority - z.priority || a.doc - z.doc);
	const assigned = new Map<number, number>();
	for (let i = 0; i < order.length; i++) {
		const rule = order[i];
		if (rule !== undefined) assigned.set(rule.doc, i + 1);
	}

	let out = "";
	for (let b = 0; b < blockRules.length; b++) {
		let rulesXml = "";
		for (const rule of blockRules[b] ?? []) {
			const p = assigned.get(rule.doc) ?? 1;
			const head = `${rule.open} priority="${p}"${rule.rest}`;
			rulesXml +=
				rule.children === ""
					? `<cfRule ${head}/>`
					: `<cfRule ${head}>${rule.children}</cfRule>`;
		}
		const tokens = blockSqref[b] ?? [];
		out += `<conditionalFormatting sqref="${tokens.join(" ")}">${rulesXml}</conditionalFormatting>`;
	}
	return out;
}

// Validate one <cfRule> and return its emit-ready pieces. Discriminates on `type`.
function buildCfRule(
	sheetName: string,
	what: string,
	raw: unknown,
	styles: StyleRegistry,
	block: number,
	doc: number,
): BuiltCfRule {
	if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
	const fail: Fail = (message) => sheetInvalid(sheetName, `${what}: ${message}`);

	const type = ownValue(raw, "type");
	if (typeof type !== "string" || !isXmlSafe(type)) {
		sheetInvalid(sheetName, `${what}.type must be an XML-safe string`);
	}
	const priority = ownValue(raw, "priority");
	if (typeof priority !== "number" || !Number.isInteger(priority) || priority < 1) {
		sheetInvalid(sheetName, `${what}.priority must be an integer ≥ 1`);
	}

	// Shared attributes across every variant.
	let rest = "";
	const stopIfTrue = ownValue(raw, "stopIfTrue");
	if (stopIfTrue !== undefined) {
		if (typeof stopIfTrue !== "boolean")
			sheetInvalid(sheetName, `${what}.stopIfTrue must be a boolean`);
		rest += ` stopIfTrue="${stopIfTrue ? 1 : 0}"`;
	}

	const finish = (open: string, extraRest: string, children: string): BuiltCfRule => ({
		block,
		doc,
		priority,
		open,
		rest: rest + extraRest,
		children,
	});

	if (type === "colorScale") {
		checkKeys(sheetName, what, raw, ["type", "priority", "stopIfTrue", "cfvo", "colors"]);
		const cfvos = cfvoXml(sheetName, what, ownValue(raw, "cfvo"));
		const colors = cfColorsXml(fail, `${what}.colors`, ownValue(raw, "colors"));
		// Excel repair-prompts on an out-of-count scale — reject (the reader drops it; shared bound).
		if (!colorScaleCountsOk(cfvos.count, colors.count)) {
			sheetInvalid(
				sheetName,
				`${what} colorScale needs 2 or 3 cfvo and the same number of colors`,
			);
		}
		return finish(
			'type="colorScale"',
			"",
			`<colorScale>${cfvos.xml}${colors.xml}</colorScale>`,
		);
	}
	if (type === "dataBar") {
		checkKeys(sheetName, what, raw, ["type", "priority", "stopIfTrue", "cfvo", "color"]);
		const cfvos = cfvoXml(sheetName, what, ownValue(raw, "cfvo"));
		if (!dataBarCountsOk(cfvos.count)) {
			sheetInvalid(sheetName, `${what} dataBar needs exactly 2 cfvo`);
		}
		const color = colorXml(
			"color",
			validateColor(fail, `${what}.color`, ownValue(raw, "color")),
		);
		return finish('type="dataBar"', "", `<dataBar>${cfvos.xml}${color}</dataBar>`);
	}
	if (type === "iconSet") {
		checkKeys(sheetName, what, raw, ["type", "priority", "stopIfTrue", "iconSet", "cfvo"]);
		let attrs = "";
		const iconSet = ownValue(raw, "iconSet");
		if (iconSet !== undefined) {
			if (typeof iconSet !== "string" || !isXmlSafe(iconSet)) {
				sheetInvalid(sheetName, `${what}.iconSet must be an XML-safe string`);
			}
			attrs += ` iconSet="${escapeAttr(iconSet)}"`;
		}
		const cfvos = cfvoXml(sheetName, what, ownValue(raw, "cfvo"));
		const iconName = typeof iconSet === "string" ? iconSet : undefined;
		if (!iconSetCountsOk(iconName, cfvos.count)) {
			sheetInvalid(
				sheetName,
				`${what} iconSet needs one cfvo per icon (${iconSetCount(iconName) ?? "3–5"})`,
			);
		}
		return finish('type="iconSet"', "", `<iconSet${attrs}>${cfvos.xml}</iconSet>`);
	}
	if (CF_HIGHLIGHT_TYPES.has(type)) {
		checkKeys(sheetName, what, raw, [
			"type",
			"priority",
			"stopIfTrue",
			"dxf",
			"operator",
			"text",
			"timePeriod",
			"rank",
			"percent",
			"bottom",
			"aboveAverage",
			"equalAverage",
			"stdDev",
			"formulas",
		]);
		const dxfId = styles.internDxf(ownValue(raw, "dxf"), `${what}.dxf`);
		let extra = "";
		const strAttr = (key: string, value: unknown): void => {
			if (value === undefined) return;
			if (typeof value !== "string" || !isXmlSafe(value)) {
				sheetInvalid(sheetName, `${what}.${key} must be an XML-safe string`);
			}
			extra += ` ${key}="${escapeAttr(value)}"`;
		};
		const intAttr = (key: string, value: unknown): void => {
			if (value === undefined) return;
			if (typeof value !== "number" || !Number.isInteger(value)) {
				sheetInvalid(sheetName, `${what}.${key} must be an integer`);
			}
			extra += ` ${key}="${value}"`;
		};
		const boolAttr = (key: string, value: unknown): void => {
			if (value === undefined) return;
			if (typeof value !== "boolean")
				sheetInvalid(sheetName, `${what}.${key} must be a boolean`);
			extra += ` ${key}="${value ? 1 : 0}"`;
		};
		strAttr("operator", ownValue(raw, "operator"));
		strAttr("text", ownValue(raw, "text"));
		strAttr("timePeriod", ownValue(raw, "timePeriod"));
		intAttr("rank", ownValue(raw, "rank"));
		boolAttr("percent", ownValue(raw, "percent"));
		boolAttr("bottom", ownValue(raw, "bottom"));
		boolAttr("aboveAverage", ownValue(raw, "aboveAverage"));
		boolAttr("equalAverage", ownValue(raw, "equalAverage"));
		intAttr("stdDev", ownValue(raw, "stdDev"));
		const children = cfFormulasXml(sheetName, what, ownValue(raw, "formulas"));
		const open = dxfId !== undefined ? `type="${type}" dxfId="${dxfId}"` : `type="${type}"`;
		return finish(open, extra, children);
	}
	sheetInvalid(
		sheetName,
		`${what}.type "${shortened(type)}" is not a known conditional-format type`,
	);
}

// <cfvo> children — thresholds for a colorScale/dataBar/iconSet. type + optional val + optional gte.
// Returns the emitted XML and the COUNT (length read once — TOCTOU-safe) so the caller can enforce
// the per-type child-count bound (decision 5).
function cfvoXml(sheetName: string, what: string, raw: unknown): { xml: string; count: number } {
	if (!Array.isArray(raw)) sheetInvalid(sheetName, `${what}.cfvo must be an array`);
	const n = raw.length;
	let out = "";
	for (let i = 0; i < n; i++) {
		const c = ownValue(raw, i) as unknown;
		const cw = `${what}.cfvo[${i}]`;
		if (!isPlainRecord(c)) sheetInvalid(sheetName, `${cw} must be an object`);
		checkKeys(sheetName, cw, c, ["type", "val", "gte"]);
		const t = ownValue(c, "type");
		if (typeof t !== "string" || !isXmlSafe(t)) {
			sheetInvalid(sheetName, `${cw}.type must be an XML-safe string`);
		}
		let attrs = ` type="${escapeAttr(t)}"`;
		const val = ownValue(c, "val");
		if (val !== undefined) {
			if (typeof val !== "string" || !isXmlSafe(val)) {
				sheetInvalid(sheetName, `${cw}.val must be an XML-safe string`);
			}
			attrs += ` val="${escapeAttr(val)}"`;
		}
		const gte = ownValue(c, "gte");
		if (gte !== undefined) {
			if (typeof gte !== "boolean") sheetInvalid(sheetName, `${cw}.gte must be a boolean`);
			attrs += ` gte="${gte ? 1 : 0}"`;
		}
		out += `<cfvo${attrs}/>`;
	}
	return { xml: out, count: n };
}

// <color> children for a colorScale. Returns the XML and the count (length read once — TOCTOU-safe).
function cfColorsXml(fail: Fail, what: string, raw: unknown): { xml: string; count: number } {
	if (!Array.isArray(raw)) fail(`${what} must be an array`);
	const n = raw.length;
	let out = "";
	for (let i = 0; i < n; i++) {
		out += colorXml("color", validateColor(fail, `${what}[${i}]`, ownValue(raw, i)));
	}
	return { xml: out, count: n };
}

// <formula> children — 0..3 operand formulas, stored form (leading = stripped), XML-safe.
function cfFormulasXml(sheetName: string, what: string, raw: unknown): string {
	if (raw === undefined) return "";
	if (!Array.isArray(raw)) sheetInvalid(sheetName, `${what}.formulas must be an array`);
	let out = "";
	let emitted = 0;
	for (let i = 0; i < raw.length; i++) {
		const f = ownValue(raw, i) as unknown;
		if (typeof f !== "string")
			sheetInvalid(sheetName, `${what}.formulas[${i}] must be a string`);
		const stored = f.startsWith("=") ? f.slice(1) : f;
		if (stored === "") continue;
		if (!isXmlSafe(stored)) {
			sheetInvalid(
				sheetName,
				`${what}.formulas[${i}] contains a character not allowed in XML`,
			);
		}
		// CT_CfRule allows at most MAX_CF_FORMULAS <formula> children — the shared bound the
		// tolerant reader also enforces (it ignores the excess; the writer refuses, F9.6).
		if (++emitted > MAX_CF_FORMULAS) {
			sheetInvalid(sheetName, `${what} has more than ${MAX_CF_FORMULAS} formulas`);
		}
		// Logical ST_Formula text, encoded once right before escaping (F11.20).
		out += `<formula>${escapeText(encodeXstring(stored))}</formula>`;
	}
	return out;
}

// ── Tables (F9.1) ──────────────────────────────────────────────────────────────────────────────

/**
 * Workbook-wide state the per-sheet table builders share: a global table-number counter (table `id`
 * and part path are workbook-global, so `id == part number`) and the set of display names claimed so
 * far (uniqueness is case-insensitive across the whole workbook, like sheet names). One instance per
 * `writeXlsx`/`streamXlsx` call, threaded through every sheet.
 */
export interface TableContext {
	/** Reserve the next global table number (1-based) — used for both `id` and the part path. */
	reserveNumber(): number;
	/** Record a display name, rejecting a case-insensitive duplicate seen on any sheet. */
	claimName(sheetName: string, name: string): void;
}

export function createTableContext(): TableContext {
	let n = 0;
	const names = new Set<string>();
	return {
		reserveNumber: () => ++n,
		claimName: (sheetName, name) => {
			const key = name.toLowerCase();
			if (names.has(key)) {
				sheetInvalid(
					sheetName,
					`table name "${shortened(name)}" is not unique across the workbook (names are compared case-insensitively)`,
				);
			}
			names.add(key);
		},
	};
}

// A table name is a defined-name-style identifier (decision 8): non-empty, ≤255, no whitespace, must
// start with a letter/underscore/backslash, and must NOT look like a cell reference or the reserved
// bare `C`/`R`. The rules themselves live in ooxml/table.ts (`tableNameProblem`), single-sourced so
// the reader normalizes exactly what the writer here rejects; this only maps each problem to a message.
function validateTableName(sheetName: string, what: string, name: string): void {
	const problem = tableNameProblem(name);
	if (problem === undefined) return;
	const messages: Record<TableNameProblem, string> = {
		empty: `${what}.name must be a non-empty string`,
		"too-long": `${what}.name exceeds ${MAX_TABLE_NAME_LEN} characters`,
		"not-xml-safe": `${what}.name contains a character not allowed in XML`,
		whitespace: `${what}.name "${shortened(name)}" must not contain whitespace`,
		"bad-start": `${what}.name "${shortened(name)}" must start with a letter, underscore, or backslash`,
		"cell-ref": `${what}.name "${name}" must not look like a cell reference`,
	};
	sheetInvalid(sheetName, messages[problem]);
}

function tableStyleInfoXml(sheetName: string, what: string, raw: unknown): string {
	if (raw === undefined) return "";
	if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what}.style must be an object`);
	checkKeys(sheetName, what, raw, [
		"name",
		"showFirstColumn",
		"showLastColumn",
		"showRowStripes",
		"showColumnStripes",
	]);
	let attrs = "";
	const styleName = ownValue(raw, "name");
	if (styleName !== undefined) {
		if (typeof styleName !== "string")
			sheetInvalid(sheetName, `${what}.style.name must be a string`);
		if (!isXmlSafe(styleName)) {
			sheetInvalid(sheetName, `${what}.style.name contains a character not allowed in XML`);
		}
		attrs += ` name="${escapeAttr(styleName)}"`;
	}
	for (const flag of [
		"showFirstColumn",
		"showLastColumn",
		"showRowStripes",
		"showColumnStripes",
	] as const) {
		const value = ownValue(raw, flag);
		if (value === undefined) continue;
		if (typeof value !== "boolean")
			sheetInvalid(sheetName, `${what}.style.${flag} must be a boolean`);
		attrs += ` ${flag}="${value ? 1 : 0}"`;
	}
	return `<tableStyleInfo${attrs}/>`;
}

/**
 * A table's caller entry for column `c`, read once (single-read TOCTOU: a Proxy array must not return a
 * different object between the check and its use). An absent entry (no `columns`, a shorter array, a
 * hole) is `undefined`, and the column takes its name from the header row. A present entry must be a
 * plain record: a class instance, a string or `null` is rejected rather than ignored, so no
 * totals-row or formula field it carries is lost without an error (F11.34).
 *
 * @example providedColumn("S", "tables[0]", undefined, 0) // undefined
 * @example providedColumn("S", "tables[0]", [{ name: "Amount" }], 0) // { name: "Amount" }
 * @example providedColumn("S", "tables[0]", [new (class { name = "Amount" })()], 0) // throws XlsxError invalid-input: sheet "S": tables[0].columns[0] must be an object
 */
export function providedColumn(
	sheetName: string,
	what: string,
	columns: readonly unknown[] | undefined,
	c: number,
): Record<string, unknown> | undefined {
	const raw = columns !== undefined ? ownValue(columns, c) : undefined;
	if (raw === undefined) return undefined;
	if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what}.columns[${c}] must be an object`);
	return raw;
}

/**
 * Validate and emit a sheet's tables (decision 8). Column names DERIVE from the header row when a
 * `headerText` resolver is supplied (the buffered writer, which has all rows); the streaming writer
 * passes `undefined` and column names come from `tables[i].columns` instead (it can't read the header
 * upfront). `ctx` assigns workbook-global ids and enforces cross-sheet name uniqueness. Returns
 * `undefined` when the sheet has no tables so an unused feature emits nothing (byte-identity).
 */
export function buildTables(
	sheetName: string,
	tables: readonly TableInfo[] | undefined,
	headerText: ((row: number, col: number) => string | undefined) | undefined,
	ctx: TableContext,
	styles: StyleRegistry,
): TablePart[] | undefined {
	if (tables === undefined) return undefined;
	if (!Array.isArray(tables)) sheetInvalid(sheetName, "tables must be an array");
	if (tables.length === 0) return undefined;
	const rects: { c1: number; r1: number; c2: number; r2: number; name: string }[] = [];
	const parts: TablePart[] = [];

	for (let i = 0; i < tables.length; i++) {
		const raw = ownValue(tables, i) as unknown;
		const what = `tables[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, [
			"name",
			"ref",
			"columns",
			"headerRow",
			"totalsRow",
			"style",
			"headerRowStyle",
			"dataStyle",
			"totalsRowStyle",
		]);

		const name = ownValue(raw, "name");
		if (typeof name !== "string") sheetInvalid(sheetName, `${what}.name must be a string`);
		validateTableName(sheetName, what, name);
		ctx.claimName(sheetName, name);

		const ref = ownValue(raw, "ref");
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const colon = ref.indexOf(":");
		const from = colon === -1 ? undefined : parseCanonicalCell(ref.slice(0, colon));
		const to = colon === -1 ? undefined : parseCanonicalCell(ref.slice(colon + 1));
		if (from === undefined || to === undefined || to.col < from.col || to.row < from.row) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" is not a canonical A1 range like "A1:C5" within Excel's grid`,
			);
		}
		const width = to.col - from.col + 1;

		const headerRow = ownValue(raw, "headerRow") !== false; // default true
		const totalsRow = ownValue(raw, "totalsRow") === true;

		const columns = ownValue(raw, "columns");
		if (columns !== undefined && !Array.isArray(columns)) {
			sheetInvalid(sheetName, `${what}.columns must be an array`);
		}
		// An empty `columns` means "derive everything from the header row"; a non-empty one must have one
		// entry per column (it supplies per-column totals metadata, aligned by position).
		if (columns !== undefined && columns.length > 0 && columns.length !== width) {
			sheetInvalid(
				sheetName,
				`${what}.columns has ${columns.length} entries but the ref width is ${width}`,
			);
		}

		// Derive each column name, and collect the totals-row/calculated formulas carried per column.
		const seen = new Set<string>();
		const columnXmls: string[] = [];
		for (let c = 0; c < width; c++) {
			const provided = providedColumn(sheetName, what, columns, c);
			let columnName: string;
			if (headerRow && headerText !== undefined) {
				const text = headerText(from.row, from.col + c);
				const provName = provided !== undefined ? ownValue(provided, "name") : undefined;
				if (text !== undefined && text !== "") {
					// Header text is authoritative — Excel requires the column name to match the header cell.
					columnName = text;
				} else if (typeof provName === "string" && provName !== "") {
					// F9.5: a non-text/blank header cell (a foreign table with numeric/empty headers) would
					// otherwise abort; fall back to the name the reader carried from `<tableColumn name>`.
					columnName = provName;
				} else {
					sheetInvalid(
						sheetName,
						`${what} header cell for column ${c + 1} must be non-empty text (table column names derive from the header row)`,
					);
				}
			} else {
				const provName = provided !== undefined ? ownValue(provided, "name") : undefined;
				if (typeof provName === "string" && provName !== "") columnName = provName;
				else if (headerRow) {
					sheetInvalid(
						sheetName,
						`${what}.columns[${c}].name is required — the streaming writer can't read the header row to derive it`,
					);
				} else columnName = `Column${c + 1}`;
			}
			// No XML-safety rejection here: a column name is ST_Xstring like its header cell (the
			// reader decodes both), so it EMITS through the same encodeXstring — the two decoded
			// views must match or Excel repair-prompts (F9.6 review fix). Dedup on the model value.
			const key = columnName.toLowerCase();
			if (seen.has(key)) {
				sheetInvalid(
					sheetName,
					`${what} has a duplicate column name "${shortened(columnName)}"`,
				);
			}
			seen.add(key);
			columnXmls.push(tableColumnXml(sheetName, what, c + 1, columnName, provided, styles));
		}

		// Two tables on one sheet may not overlap (Excel repairs overlapping tables). Few tables per
		// sheet, so an O(n²) pairwise check is fine.
		for (const r of rects) {
			if (from.col <= r.c2 && r.c1 <= to.col && from.row <= r.r2 && r.r1 <= to.row) {
				sheetInvalid(sheetName, `tables "${r.name}" and "${name}" overlap`);
			}
		}
		rects.push({ c1: from.col, r1: from.row, c2: to.col, r2: to.row, name });

		// A totals row is the last row and must have at least one row above it — a single-row ref can't
		// carry one (it would drive the auto-filter range negative). Reject typed, never bare-throw.
		if (totalsRow && to.row === from.row) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" has a totals row but only one row — a totals row needs at least one row above it`,
			);
		}
		const number = ctx.reserveNumber();
		// The auto-filter exists only for a table WITH a header (there's nothing to filter without header
		// labels — Excel/openpyxl omit it otherwise); it covers the header + data rows, minus any totals row.
		let autoFilter = "";
		if (headerRow) {
			const filterEnd = totalsRow ? to.row - 1 : to.row;
			autoFilter = `<autoFilter ref="${formatRef({ col: from.col, row: from.row })}:${formatRef({ col: to.col, row: filterEnd })}"/>`;
		}
		const headerAttr = headerRow ? "" : ' headerRowCount="0"';
		const totalsAttr = totalsRow ? ' totalsRowCount="1"' : ' totalsRowShown="0"';
		// Table-wide highlight dxfs (F9.3 retrofit) — interned into the shared <dxfs> table.
		const tableDxf = tableDxfAttrs(styles, what, raw);
		const styleXml = tableStyleInfoXml(sheetName, what, ownValue(raw, "style"));
		// name/displayName are ST_Xstring — encode so a legal name that LOOKS like an escape
		// (`_x0041_`) reads back as itself in Excel/openpyxl (both decode table names). A name
		// with nothing to protect is unchanged (byte-identity).
		const storedName = encodeXstring(name);
		const xml =
			`${XML_DECL}\n<table xmlns="${NS_MAIN}" id="${number}" name="${escapeAttr(storedName)}" displayName="${escapeAttr(storedName)}" ref="${ref}"${headerAttr}${totalsAttr}${tableDxf}>` +
			`${autoFilter}<tableColumns count="${width}">${columnXmls.join("")}</tableColumns>${styleXml}</table>`;
		parts.push({ number, xml });
	}
	return parts;
}

// Intern a table's or column's highlight dxfs (headerRow/data/totals) into styles.xml's shared <dxfs>
// table and return the `*DxfId` attributes (F9.3 retrofit). Absent styles emit nothing (byte-identity).
function tableDxfAttrs(styles: StyleRegistry, what: string, raw: Record<string, unknown>): string {
	let attrs = "";
	const header = styles.internDxf(ownValue(raw, "headerRowStyle"), `${what}.headerRowStyle`);
	if (header !== undefined) attrs += ` headerRowDxfId="${header}"`;
	const data = styles.internDxf(ownValue(raw, "dataStyle"), `${what}.dataStyle`);
	if (data !== undefined) attrs += ` dataDxfId="${data}"`;
	const totals = styles.internDxf(ownValue(raw, "totalsRowStyle"), `${what}.totalsRowStyle`);
	if (totals !== undefined) attrs += ` totalsRowDxfId="${totals}"`;
	return attrs;
}

// One <tableColumn>. Names are validated by the caller; the optional totals-row label/function and the
// totals/calculated formulas are carried from the caller's TableColumn verbatim (never evaluated).
// @name and @totalsRowLabel are ST_Xstring (like the header cell they mirror) — encoded, so the
// decoded views agree in Excel; the function/formulas are NOT xstrings and stay strict.
function tableColumnXml(
	sheetName: string,
	what: string,
	id: number,
	name: string,
	provided: Record<string, unknown> | undefined,
	styles: StyleRegistry,
): string {
	let attrs = ` id="${id}" name="${escapeAttr(encodeXstring(name))}"`;
	let children = "";
	if (provided !== undefined) {
		checkKeys(sheetName, `${what} column`, provided, [
			"name",
			"totalsRowLabel",
			"totalsRowFunction",
			"totalsRowFormula",
			"calculatedColumnFormula",
			"headerRowStyle",
			"dataStyle",
			"totalsRowStyle",
		]);
		// Per-column highlight dxfs (F9.3 retrofit) — interned into the shared <dxfs>.
		attrs += tableDxfAttrs(styles, `${what} column ${id}`, provided);
		const label = ownValue(provided, "totalsRowLabel");
		if (label !== undefined) {
			if (typeof label !== "string")
				sheetInvalid(sheetName, `${what} totalsRowLabel must be a string`);
			attrs += ` totalsRowLabel="${escapeAttr(encodeXstring(label))}"`;
		}
		const fn = ownValue(provided, "totalsRowFunction");
		if (fn !== undefined)
			attrs += ` totalsRowFunction="${escapeAttr(requireXmlString(sheetName, `${what} totalsRowFunction`, fn))}"`;
		// Both formulas are logical ST_Formula (⊂ ST_Xstring) text: validated, then encoded once
		// right before escaping (F11.20) — the same convention the table name/label already use.
		const totalsFormula = ownValue(provided, "totalsRowFormula");
		if (totalsFormula !== undefined) {
			const logical = requireXmlString(sheetName, `${what} totalsRowFormula`, totalsFormula);
			children += `<totalsRowFormula>${escapeText(encodeXstring(logical))}</totalsRowFormula>`;
		}
		const calcFormula = ownValue(provided, "calculatedColumnFormula");
		if (calcFormula !== undefined) {
			const logical = requireXmlString(
				sheetName,
				`${what} calculatedColumnFormula`,
				calcFormula,
			);
			children += `<calculatedColumnFormula>${escapeText(encodeXstring(logical))}</calculatedColumnFormula>`;
		}
	}
	return children === ""
		? `<tableColumn${attrs}/>`
		: `<tableColumn${attrs}>${children}</tableColumn>`;
}

function requireXmlString(sheetName: string, what: string, value: unknown): string {
	if (typeof value !== "string") sheetInvalid(sheetName, `${what} must be a string`);
	if (!isXmlSafe(value))
		sheetInvalid(sheetName, `${what} contains a character not allowed in XML`);
	return value;
}

// ── Pictures (F6.3) ────────────────────────────────────────────────────────────────────────────
// MAX_EMU and the mime allowlist come from ooxml/drawing.ts — the SAME constants the tolerant
// reader clamps/derives with, so whatever the reader returns is writable here (shared bounds).
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

/** The drawing part + its rels part for one sheet's pictures. */
interface DrawingParts {
	readonly drawingXml: string;
	readonly drawingRelsXml: string;
}

// Check one size/offset number (measured in EMUs, Excel's tiny drawing unit). It must be a whole
// number from 0 up to the limit; if it's missing we treat it as 0, anything else is rejected.
function emuValue(sheetName: string, what: string, raw: unknown): number {
	if (raw === undefined) return 0;
	if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0 || raw > MAX_EMU) {
		sheetInvalid(sheetName, `${what} must be an integer EMU in 0..${MAX_EMU}`);
	}
	return raw;
}

// A checked corner of a picture: which cell it sits at (1-based column/row) and how far into that
// cell it starts (the offsets).
interface Point {
	readonly col: number;
	readonly row: number;
	readonly colOff: number;
	readonly rowOff: number;
}

// Check one corner the caller gave us: the cell must be inside the sheet's grid and the offsets must
// be valid sizes. Throws a clear error naming the bad field, otherwise returns the clean values.
function validatePoint(sheetName: string, what: string, raw: unknown): Point {
	if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
	checkKeys(sheetName, what, raw, ["col", "row", "colOff", "rowOff"]);
	const col = ownValue(raw, "col");
	const row = ownValue(raw, "row");
	if (typeof col !== "number" || !Number.isInteger(col) || col < 1 || col > MAX_COL) {
		sheetInvalid(sheetName, `${what}.col must be an integer column in 1..${MAX_COL}`);
	}
	if (typeof row !== "number" || !Number.isInteger(row) || row < 1 || row > MAX_ROW) {
		sheetInvalid(sheetName, `${what}.row must be an integer row in 1..${MAX_ROW}`);
	}
	return {
		col,
		row,
		colOff: emuValue(sheetName, `${what}.colOff`, ownValue(raw, "colOff")),
		rowOff: emuValue(sheetName, `${what}.rowOff`, ownValue(raw, "rowOff")),
	};
}

// Write one corner as XML. Our numbers count columns/rows from 1, but the file counts from 0, so we
// subtract one here.
function pointXml(tag: string, p: Point): string {
	return `<xdr:${tag}><xdr:col>${p.col - 1}</xdr:col><xdr:colOff>${p.colOff}</xdr:colOff><xdr:row>${p.row - 1}</xdr:row><xdr:rowOff>${p.rowOff}</xdr:rowOff></xdr:${tag}>`;
}

/**
 * Turn a sheet's pictures into the two XML pieces the file needs: the drawing (where each picture
 * sits) and its relationships (which image file each picture points to). Along the way it checks
 * every picture and stashes its bytes in the shared media store so duplicates are written once.
 * Returns nothing when the sheet has no pictures. Anything invalid throws an error that names the
 * sheet and the picture's position; each field the caller gave is read exactly once.
 */
function imageParts(
	sheetName: string,
	images: readonly SheetImage[] | undefined,
	media: MediaRegistry,
): DrawingParts | undefined {
	if (images === undefined) return undefined;
	if (!Array.isArray(images)) sheetInvalid(sheetName, "images must be an array");
	if (images.length === 0) return undefined;

	const anchors: string[] = [];
	const rels: string[] = [];
	for (let i = 0; i < images.length; i++) {
		const raw = ownValue(images, i) as unknown;
		const what = `images[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["anchor", "bytes", "mime", "name"]);

		// bytes — read the reference ONCE (a getter must not swap the buffer between here and packing).
		const bytes = ownValue(raw, "bytes");
		if (!(bytes instanceof Uint8Array) || bytes.length === 0) {
			sheetInvalid(sheetName, `${what}.bytes must be a non-empty Uint8Array`);
		}
		// mime — allowlist only; the extension is derived from the mime, never sniffed from the data.
		// Use hasOwn, not a bare lookup: a bare `map[mime]` walks the prototype chain, so a mime like
		// "constructor" or "toString" would resolve to an inherited value and slip past the gate.
		const mime = ownValue(raw, "mime");
		const ext =
			typeof mime === "string" && Object.hasOwn(MEDIA_MIME_TO_EXT, mime)
				? MEDIA_MIME_TO_EXT[mime]
				: undefined;
		if (ext === undefined) {
			// Enumerate the allowlist from the map itself so this message can never go stale.
			sheetInvalid(
				sheetName,
				`${what}.mime must be one of ${Object.keys(MEDIA_MIME_TO_EXT).join(", ")}`,
			);
		}
		// name — optional; defaults to a deterministic "Image N".
		const rawName = ownValue(raw, "name");
		if (rawName !== undefined && (typeof rawName !== "string" || !isXmlSafe(rawName))) {
			sheetInvalid(sheetName, `${what}.name must be an XML-safe string`);
		}
		const name = rawName ?? `Image ${i + 1}`;

		// anchor — exactly one of `to` (two-cell) or `ext` (one-cell).
		const anchor = ownValue(raw, "anchor");
		if (!isPlainRecord(anchor)) sheetInvalid(sheetName, `${what}.anchor must be an object`);
		checkKeys(sheetName, `${what}.anchor`, anchor, ["from", "to", "ext", "editAs"]);
		const from = validatePoint(sheetName, `${what}.anchor.from`, ownValue(anchor, "from"));
		// Read `to`/`ext` ONCE each (single-read TOCTOU): a getter must not present one shape to the
		// XOR check and another to emission.
		const rawTo = ownValue(anchor, "to");
		const rawExt = ownValue(anchor, "ext");
		const hasTo = rawTo !== undefined;
		const hasExt = rawExt !== undefined;
		if (hasTo === hasExt) {
			sheetInvalid(
				sheetName,
				`${what}.anchor must have exactly one of "to" (two-cell) or "ext" (one-cell)`,
			);
		}
		const editAs = ownValue(anchor, "editAs");
		if (
			editAs !== undefined &&
			editAs !== "twoCell" &&
			editAs !== "oneCell" &&
			editAs !== "absolute"
		) {
			sheetInvalid(sheetName, `${what}.anchor.editAs must be twoCell, oneCell, or absolute`);
		}

		const rid = `rId${i + 1}`;
		const mediaNumber = media.intern(bytes, ext);
		rels.push(
			`<Relationship Id="${rid}" Type="${NS_REL}/image" Target="../media/image${mediaNumber}.${ext}"/>`,
		);
		const pic =
			`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="${escapeAttr(name)}"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
			`<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
			`<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
		if (hasTo) {
			const to = validatePoint(sheetName, `${what}.anchor.to`, rawTo);
			const editAsAttr = editAs !== undefined ? ` editAs="${editAs}"` : "";
			anchors.push(
				`<xdr:twoCellAnchor${editAsAttr}>${pointXml("from", from)}${pointXml("to", to)}${pic}<xdr:clientData/></xdr:twoCellAnchor>`,
			);
		} else {
			if (!isPlainRecord(rawExt))
				sheetInvalid(sheetName, `${what}.anchor.ext must be an object`);
			checkKeys(sheetName, `${what}.anchor.ext`, rawExt, ["cx", "cy"]);
			const cx = emuValue(sheetName, `${what}.anchor.ext.cx`, ownValue(rawExt, "cx"));
			const cy = emuValue(sheetName, `${what}.anchor.ext.cy`, ownValue(rawExt, "cy"));
			anchors.push(
				`<xdr:oneCellAnchor>${pointXml("from", from)}<xdr:ext cx="${cx}" cy="${cy}"/>${pic}<xdr:clientData/></xdr:oneCellAnchor>`,
			);
		}
	}

	const drawingXml = `${XML_DECL}\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_REL}">${anchors.join("")}</xdr:wsDr>`;
	const drawingRelsXml = `${XML_DECL}\n<Relationships xmlns="${NS_PKG_REL}">${rels.join("")}</Relationships>`;
	return { drawingXml, drawingRelsXml };
}

/** A single-sourced sheet relationship — one <Relationship> line in the sheet's rels part. */
export interface SheetRel {
	/** Full relationship type URI (e.g. `${NS_REL}/comments`). */
	readonly type: string;
	/** The Target value: an external URL, or an internal part path relative to the worksheet. */
	readonly target: string;
	/** TargetMode="External" — set for hyperlink URLs, clear for internal parts. */
	readonly external: boolean;
}

/**
 * The auxiliary parts a worksheet carries beside its own body: a rels part (present when the sheet
 * has any relationships) and, for comments, the paired comments + VML drawing parts. Shared by the
 * buffered {@link WorksheetResult} and the streaming {@link StreamedWorksheet} so the two writers
 * can't drift; `sheetSideParts` (parts.ts) turns these into the actual OPC parts, one place owning
 * the part names.
 */
export interface SheetSideParts {
	/**
	 * The sheet's relationships in rId order (rId = index + 1) — non-empty means it needs a rels
	 * part. Hyperlinks come first so a hyperlinks-only sheet keeps its exact pre-F5.2 rels bytes.
	 */
	readonly rels: readonly SheetRel[];
	/** `xl/commentsN.xml` content — present iff the sheet has comments (paired with {@link vmlXml}). */
	readonly commentsXml?: string;
	/** `xl/drawings/vmlDrawingN.vml` content — the legacy drawing that makes comments render. */
	readonly vmlXml?: string;
	/** `xl/drawings/drawingN.xml` content — present iff the sheet has pictures (F6.3). */
	readonly drawingXml?: string;
	/** `xl/drawings/_rels/drawingN.xml.rels` — the picture blip → media relationships. */
	readonly drawingRelsXml?: string;
	/** `xl/tables/tableN.xml` parts owned by this sheet (F9.1) — present iff the sheet has tables. */
	readonly tables?: readonly TablePart[];
	/**
	 * The sheet's validated autoFilter range (F10.2), when it has one — NOT an OPC part but the range the
	 * WORKBOOK writer turns into the paired hidden `_xlnm._FilterDatabase` defined name (sheet-scoped to
	 * this sheet). Absent when the sheet declares no autoFilter. Read once here so the range is validated
	 * a single time and reused for both the `<autoFilter>` element and the synthesized name.
	 */
	readonly autoFilterRef?: string;
}

/** One emitted table part: its workbook-global number (→ `xl/tables/table{number}.xml`) and its XML. */
export interface TablePart {
	readonly number: number;
	readonly xml: string;
}

export interface WorksheetResult extends SheetSideParts {
	readonly xml: string;
}

/**
 * Build a sheet's relationships and the body plumbing that references them — the block both writers
 * had copied verbatim. Hyperlink targets take rId1..rIdN (matching the r:ids hyperlinksXml already
 * emitted), then the comments + vmlDrawing pair; `legacyDrawing` is the body element pointing at the
 * VML part, and `nsR` is the `xmlns:r` the root needs when the body references any rId. Single owner
 * of the per-sheet rel ORDER, rId allocation, and the comment/VML part-path convention.
 */
function sheetRelPlumbing(
	sheetIndex: number,
	hyperlinkTargets: readonly string[],
	hasComments: boolean,
	hasDrawing: boolean,
	tableNumbers: readonly number[],
): { rels: SheetRel[]; drawing: string; legacyDrawing: string; tableParts: string; nsR: string } {
	const rels: SheetRel[] = hyperlinkTargets.map((target) => ({
		type: `${NS_REL}/hyperlink`,
		target,
		external: true,
	}));
	// The drawingML part (F6.3) is referenced by <drawing r:id>, which sits before <legacyDrawing> in
	// the CT_Worksheet sequence — so it takes its rId before the comments/VML pair.
	let drawing = "";
	if (hasDrawing) {
		rels.push({
			type: `${NS_REL}/drawing`,
			target: `../drawings/drawing${sheetIndex + 1}.xml`,
			external: false,
		});
		drawing = `<drawing r:id="rId${rels.length}"/>`;
	}
	let legacyDrawing = "";
	if (hasComments) {
		// The comments part is located by rel TYPE (no r:id in the sheet body); the vmlDrawing IS
		// referenced by <legacyDrawing r:id>. Targets are relative to the worksheet's own directory.
		rels.push({
			type: `${NS_REL}/comments`,
			target: `../comments${sheetIndex + 1}.xml`,
			external: false,
		});
		rels.push({
			type: `${NS_REL}/vmlDrawing`,
			target: `../drawings/vmlDrawing${sheetIndex + 1}.vml`,
			external: false,
		});
		legacyDrawing = `<legacyDrawing r:id="rId${rels.length}"/>`; // the vmlDrawing rel, just pushed
	}
	// Tables are the LAST child of CT_Worksheet (`<tableParts>` after `<legacyDrawing>`), so their rels
	// take the highest rIds. One `<tablePart r:id>` per table; targets are workbook-global part paths.
	let tableParts = "";
	if (tableNumbers.length > 0) {
		const partRefs = tableNumbers
			.map((num) => {
				rels.push({
					type: `${NS_REL}/table`,
					target: `../tables/table${num}.xml`,
					external: false,
				});
				return `<tablePart r:id="rId${rels.length}"/>`;
			})
			.join("");
		tableParts = `<tableParts count="${tableNumbers.length}">${partRefs}</tableParts>`;
	}
	// xmlns:r is declared whenever the body references an rId — a hyperlink, drawing, legacyDrawing, or
	// tablePart. A sheet using none keeps the exact pre-F4.6 root element.
	const nsR =
		hyperlinkTargets.length > 0 || drawing !== "" || legacyDrawing !== "" || tableParts !== ""
			? ` xmlns:r="${NS_REL}"`
			: "";
	return { rels, drawing, legacyDrawing, tableParts, nsR };
}

/**
 * Build the worksheet XML for one sheet, interning cell styles into the shared registry.
 * `rows` is the reference captured during workbook validation, also used for table-header lookup.
 * `sheetIndex` (0-based) names the sheet's own extra parts (comments, VML) so their rel targets and
 * package paths agree.
 */
export function worksheetXml(
	sheet: Omit<SheetInput, "rows">,
	rows: SheetInput["rows"],
	sheetIndex: number,
	date1904: boolean,
	styles: StyleRegistry,
	media: MediaRegistry,
	tableCtx: TableContext,
	sheetName: string,
): WorksheetResult {
	// Geometry and metadata validate up front (geometry also contributes rows below): a bad
	// column/row/freeze/merge/hyperlink/comment/image spec must surface before any cell work.
	const cols = colsXml(sheetName, ownValue(sheet, "columns"));
	const rowAttrs = rowAttrsMap(sheetName, ownValue(sheet, "rowProperties"));
	const sheetViews = sheetViewsXml(sheetName, ownValue(sheet, "freeze"));
	const autoFilter = autoFilterXml(sheetName, ownValue(sheet, "autoFilter"));
	const sheetProtection = sheetProtectionXml(sheetName, ownValue(sheet, "protection"));
	const printSetup = printSetupXml(sheetName, sheet);
	const mergeCells = mergeCellsXml(sheetName, ownValue(sheet, "merges"));
	const conditionalFormatting = conditionalFormattingXml(
		sheetName,
		ownValue(sheet, "conditionalFormatting"),
		styles,
	);
	const dataValidations = dataValidationsXml(sheetName, ownValue(sheet, "dataValidations"));
	const links = hyperlinksXml(sheetName, ownValue(sheet, "hyperlinks"));
	const comments = commentsParts(sheetName, ownValue(sheet, "comments"));
	const pictures = imageParts(sheetName, ownValue(sheet, "images"), media);
	// Only table-header coordinates need to survive both metadata and cell emission. Cache by
	// position, not object identity: callers can reuse one StyledCell at distinct coordinates.
	const headerCells = new Map<
		number,
		Map<number, { readonly cell: CapturedCell; readonly text: string | undefined }>
	>();
	const headerRows = new Map<number, SheetInput["rows"][number]>();
	const tables = buildTables(
		sheetName,
		ownValue(sheet, "tables"),
		(row, col) => {
			let capturedRow = headerCells.get(row);
			const captured = capturedRow?.get(col);
			if (captured !== undefined) return captured.text;
			if (!headerRows.has(row - 1)) headerRows.set(row - 1, ownValue(rows, row - 1));
			const rowCells = headerRows.get(row - 1);
			const input: CellInput = Array.isArray(rowCells)
				? ownValue(rowCells, col - 1)
				: undefined;
			// Preserve header policy: bare strings and plain records with a string value (including
			// cached formula strings) infer names; other cells may use a provided column-name fallback.
			const canInfer = typeof input === "string" || isPlainRecord(input);
			const cell = splitInput(col, row, input);
			const text = canInfer && typeof cell.value === "string" ? cell.value : undefined;
			if (capturedRow === undefined) {
				capturedRow = new Map();
				headerCells.set(row, capturedRow);
			}
			capturedRow.set(col, { cell, text });
			return text;
		},
		tableCtx,
		styles,
	);

	// 0 means "unset" — no populated cell seen yet (columns/rows are 1-based, so 0 is a safe sentinel).
	let minRow = 0;
	let maxRow = 0;
	let minCol = 0;
	let maxCol = 0;
	// (rowNum, xml) pairs: cell rows arrive in ascending order; property-only rows are merged in
	// afterwards, then the whole set is sorted so <sheetData> stays ascending.
	const rowXmls: [number, string][] = [];

	// A workbook can't outgrow Excel's grid: refs past XFD1048576 make Excel refuse the file, and
	// (mechanically) `rows.length` drives this loop — an absurd length would spin for hours.
	if (rows.length > MAX_ROW) {
		throw new XlsxError("invalid-input", `a sheet cannot have more than ${MAX_ROW} rows`);
	}

	const listings: RowListings = new WeakMap();
	for (let r = 0; r < rows.length; r++) {
		const cells = headerRows.has(r) ? headerRows.get(r) : ownValue(rows, r);
		// A missing row (array hole / undefined) is an empty row — skip it. Anything else that isn't
		// an array (a string, a number, null, an object) would otherwise be iterated as if it were a
		// row: a string "abc" would explode into three character cells. Reject it instead of silently
		// mangling the data — this also turns a null row into a clean error rather than a TypeError.
		if (cells === undefined) continue;
		if (!Array.isArray(cells)) {
			throw new XlsxError(
				"invalid-input",
				`sheet row ${r + 1}: a row must be an array of cell values`,
			);
		}
		const length = cells.length;
		if (length === 0) continue;
		if (length > MAX_COL) {
			throw new XlsxError(
				"invalid-input",
				`sheet row ${r + 1}: a row cannot have more than ${MAX_COL} cells`,
			);
		}
		const rowNum = r + 1;
		const capturedRow = headerCells.get(rowNum);
		const captured =
			capturedRow === undefined
				? undefined
				: new Set([...capturedRow.keys()].map((col) => col - 1));
		const cellXmls: string[] = [];
		forEachRowColumn(
			cells,
			length,
			(c, own) => {
				const colNum = c + 1;
				const cell =
					capturedRow?.get(colNum)?.cell ??
					splitInput(colNum, rowNum, own ? cells[c] : undefined);
				const rendered = renderCell(colNum, rowNum, cell, date1904, styles);
				if (rendered === undefined) return;
				if (minRow === 0 || rowNum < minRow) minRow = rowNum;
				if (rowNum > maxRow) maxRow = rowNum;
				if (minCol === 0 || colNum < minCol) minCol = colNum;
				if (colNum > maxCol) maxCol = colNum;
				cellXmls.push(rendered);
			},
			captured,
			listings,
		);
		if (cellXmls.length > 0) {
			const attrs = rowAttrs.get(rowNum) ?? "";
			rowAttrs.delete(rowNum); // consumed — whatever remains becomes a property-only row
			rowXmls.push([rowNum, `<row r="${rowNum}"${attrs}>${cellXmls.join("")}</row>`]);
		}
	}

	// Rows that carry height/hidden but no cells still exist in the file, as cell-less <row>
	// elements. They do not extend the dimension (Excel's dimension covers content, not geometry).
	for (const [rowNum, attrs] of rowAttrs) {
		rowXmls.push([rowNum, `<row r="${rowNum}"${attrs}/>`]);
	}
	rowXmls.sort((a, b) => a[0] - b[0]);

	// Bounding box of the populated cells, in A1 notation. An entirely empty sheet is "A1" (Excel's
	// convention); a single cell collapses to that one ref rather than a degenerate "X:X" range.
	const dimension =
		minRow === 0
			? "A1"
			: minRow === maxRow && minCol === maxCol
				? formatRef({ col: minCol, row: minRow })
				: `${formatRef({ col: minCol, row: minRow })}:${formatRef({ col: maxCol, row: maxRow })}`;

	// Relationships + the body plumbing that references them (shared with the streaming writer).
	const { rels, drawing, legacyDrawing, tableParts, nsR } = sheetRelPlumbing(
		sheetIndex,
		links.targets,
		comments !== undefined,
		pictures !== undefined,
		tables !== undefined ? tables.map((t) => t.number) : [],
	);

	// Schema order within <worksheet>: dimension, sheetViews, cols, sheetData, sheetProtection,
	// autoFilter, mergeCells, conditionalFormatting, dataValidations, hyperlinks, drawing, legacyDrawing,
	// tableParts (CT_Worksheet sequence — sheetProtection then autoFilter sit between sheetData and
	// mergeCells; conditionalFormatting then dataValidations sit between mergeCells and hyperlinks;
	// tableParts is last). Every optional block is an empty string when unused, so a sheet using none
	// emits the exact pre-F4.5/…/F10.3 bytes.
	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"${nsR}><dimension ref="${dimension}"/>${sheetViews}${cols}<sheetData>${rowXmls
		.map(([, x]) => x)
		.join(
			"",
		)}</sheetData>${sheetProtection}${autoFilter.xml}${mergeCells}${conditionalFormatting}${dataValidations}${links.xml}${printSetup}${drawing}${legacyDrawing}${tableParts}</worksheet>`;
	// Keep absent side parts absent even if Object.prototype changes before the parts are emitted.
	const result: WorksheetResult = {
		xml,
		rels,
		...(comments !== undefined
			? { commentsXml: comments.commentsXml, vmlXml: comments.vmlXml }
			: {}),
		...(pictures !== undefined
			? { drawingXml: pictures.drawingXml, drawingRelsXml: pictures.drawingRelsXml }
			: {}),
		...(tables !== undefined ? { tables } : {}),
		...(autoFilter.ref !== undefined ? { autoFilterRef: autoFilter.ref } : {}),
	};
	Object.setPrototypeOf(result, null);
	return result;
}

// ── Streaming worksheet (F5.1) ─────────────────────────────────────────────────────────────────
// The constant-memory sibling of worksheetXml: geometry and metadata are computed upfront (the same
// helpers/validators), the header and footer are strings, and only the rows flow — each rendered
// through renderCell and interning styles as it passes. A streamed sheet OMITS <dimension> (its
// bounding box is unknowable before the rows arrive; the element is optional and Excel/openpyxl
// recompute it). Everything else — element order, style interning, metadata — matches the buffered
// writer, so the two produce reader-equivalent output.

/** A streamed worksheet: its XML as a chunk generator, plus the upfront-known rels/comment parts. */
export interface StreamedWorksheet extends SheetSideParts {
	readonly chunks: AsyncGenerator<Uint8Array>;
}

/**
 * Render one streamed row. Its number is the 1-based position in the stream. Mirrors the buffered
 * per-row logic: an empty row with no properties emits nothing; a property-only row emits `<row r/>`;
 * otherwise `<row r>cells</row>`. Interns styles into the shared registry as it renders, and visits
 * only the row's populated columns (F11.54); `listings` keeps key listings across one sheet's rows.
 *
 * @example renderStreamRow("S", 3, undefined, new Map(), false, styles) // ""
 * @example renderStreamRow("S", 3, [], new Map([[3, ' ht="20" customHeight="1"']]), false, styles) // '<row r="3" ht="20" customHeight="1"/>'
 * @example renderStreamRow("S", 3, [, 7], new Map(), false, styles) // '<row r="3"><c r="B3"><v>7</v></c></row>'
 * @example renderStreamRow("S", 3, "abc", new Map(), false, styles) // throws XlsxError invalid-input: a row must be an array of cell values
 * @example renderStreamRow("S", 3, new Array(16385), new Map(), false, styles) // throws XlsxError invalid-input: more than 16384 cells
 */
export function renderStreamRow(
	sheetName: string,
	rowNum: number,
	cells: readonly CellInput[] | undefined,
	rowAttrs: Map<number, string>,
	date1904: boolean,
	styles: StyleRegistry,
	listings?: RowListings,
): string {
	const attrs = rowAttrs.get(rowNum) ?? "";
	rowAttrs.delete(rowNum); // consumed — whatever remains becomes a trailing property-only row
	if (cells === undefined) return attrs !== "" ? `<row r="${rowNum}"${attrs}/>` : "";
	if (!Array.isArray(cells)) {
		throw new XlsxError(
			"invalid-input",
			`sheet "${sheetName}" row ${rowNum}: a row must be an array of cell values`,
		);
	}
	const length = cells.length;
	if (length > MAX_COL) {
		throw new XlsxError(
			"invalid-input",
			`sheet "${sheetName}" row ${rowNum}: a row cannot have more than ${MAX_COL} cells`,
		);
	}
	const cellXmls: string[] = [];
	forEachRowColumn(
		cells,
		length,
		(c, own) => {
			const cell = splitInput(c + 1, rowNum, own ? cells[c] : undefined);
			const rendered = renderCell(c + 1, rowNum, cell, date1904, styles);
			if (rendered !== undefined) cellXmls.push(rendered);
		},
		undefined,
		listings,
	);
	if (cellXmls.length === 0) return attrs !== "" ? `<row r="${rowNum}"${attrs}/>` : "";
	return `<row r="${rowNum}"${attrs}>${cellXmls.join("")}</row>`;
}

// Roughly this many bytes of row XML accumulate before a chunk is flushed — batching keeps the
// compressor fed with substantial writes instead of one tiny write per row.
const ROW_CHUNK_BYTES = 65536;

// Native array iteration resolves holes through the prototype chain. Preserve array positions
// with own reads when the caller uses that iterator. Custom iterable protocols remain supported.
function* ownRows(rows: SheetInput["rows"]): Generator<SheetInput["rows"][number]> {
	for (let i = 0; i < rows.length; i++) yield ownValue(rows, i);
}

function streamRowSource(rows: StreamSheetInput["rows"]): StreamSheetInput["rows"] {
	const source = rows;
	if (!Array.isArray(rows)) return rows;
	// Capture the selected protocol once, as for-await would. Returning the original array
	// after inspecting its method would invoke a caller's iterator getter a second time.
	if (Symbol.asyncIterator in source) {
		const iterate = source[Symbol.asyncIterator];
		if (iterate !== undefined && iterate !== null) {
			return { [Symbol.asyncIterator]: () => iterate.call(rows) };
		}
	}
	const iterate = rows[Symbol.iterator];
	return iterate === Array.prototype[Symbol.iterator]
		? ownRows(rows)
		: { [Symbol.iterator]: () => iterate.call(rows) };
}

async function* streamRowChunks(
	sheetName: string,
	rows: StreamSheetInput["rows"],
	header: string,
	footer: string,
	rowAttrs: Map<number, string>,
	date1904: boolean,
	styles: StyleRegistry,
): AsyncGenerator<Uint8Array> {
	yield encoder.encode(header);
	let rowNum = 0;
	const listings: RowListings = new WeakMap();
	let buf = "";
	// `for await` iterates both a sync iterable (an array) and an async one (a DB cursor).
	for await (const cells of streamRowSource(rows)) {
		rowNum++;
		if (rowNum > MAX_ROW) {
			throw new XlsxError(
				"invalid-input",
				`sheet "${sheetName}": a sheet cannot have more than ${MAX_ROW} rows`,
			);
		}
		buf += renderStreamRow(sheetName, rowNum, cells, rowAttrs, date1904, styles, listings);
		if (buf.length >= ROW_CHUNK_BYTES) {
			yield encoder.encode(buf);
			buf = "";
		}
	}
	// Row properties addressed past the last streamed row still exist in the file, as cell-less
	// <row> elements — exactly what the buffered writer emits for its leftover rowAttrs. They are
	// all beyond the streamed positions (each streamed row consumed its own entry), so appending
	// them sorted keeps the rows ascending; the map is upfront input, so memory stays constant.
	for (const [r, attrs] of [...rowAttrs].sort((a, b) => a[0] - b[0])) {
		buf += `<row r="${r}"${attrs}/>`;
	}
	if (buf !== "") yield encoder.encode(buf);
	yield encoder.encode(footer);
}

/**
 * Prepare a streamed worksheet: validate geometry/metadata and build the header/footer + rel/comment
 * parts upfront, returning a chunk generator that streams the rows on demand. Styles intern into the
 * shared registry as the generator is consumed, so styles.xml must be emitted only after it drains.
 * `rows` is the iterable captured during workbook validation; preparation never consumes it.
 */
export function streamWorksheet(
	sheet: Omit<StreamSheetInput, "rows">,
	rows: StreamSheetInput["rows"],
	sheetIndex: number,
	date1904: boolean,
	styles: StyleRegistry,
	media: MediaRegistry,
	tableCtx: TableContext,
	sheetName: string,
): StreamedWorksheet {
	const cols = colsXml(sheetName, ownValue(sheet, "columns"));
	const rowAttrs = rowAttrsMap(sheetName, ownValue(sheet, "rowProperties"));
	const sheetViews = sheetViewsXml(sheetName, ownValue(sheet, "freeze"));
	const autoFilter = autoFilterXml(sheetName, ownValue(sheet, "autoFilter"));
	const sheetProtection = sheetProtectionXml(sheetName, ownValue(sheet, "protection"));
	const printSetup = printSetupXml(sheetName, sheet);
	const mergeCells = mergeCellsXml(sheetName, ownValue(sheet, "merges"));
	const conditionalFormatting = conditionalFormattingXml(
		sheetName,
		ownValue(sheet, "conditionalFormatting"),
		styles,
	);
	const dataValidations = dataValidationsXml(sheetName, ownValue(sheet, "dataValidations"));
	const links = hyperlinksXml(sheetName, ownValue(sheet, "hyperlinks"));
	const comments = commentsParts(sheetName, ownValue(sheet, "comments"));
	const pictures = imageParts(sheetName, ownValue(sheet, "images"), media);
	// The footer (which carries <tableParts>) is built BEFORE the rows stream, so the header row isn't
	// available — column names come from `tables[i].columns` here (no header resolver).
	const tables = buildTables(sheetName, ownValue(sheet, "tables"), undefined, tableCtx, styles);

	// Relationships + the body plumbing that references them — the SAME builder the buffered writer
	// uses, so the two can't drift on rel order, rId allocation, or the drawing/legacyDrawing refs.
	const { rels, drawing, legacyDrawing, tableParts, nsR } = sheetRelPlumbing(
		sheetIndex,
		links.targets,
		comments !== undefined,
		pictures !== undefined,
		tables !== undefined ? tables.map((t) => t.number) : [],
	);

	const header = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"${nsR}>${sheetViews}${cols}<sheetData>`;
	const footer = `</sheetData>${sheetProtection}${autoFilter.xml}${mergeCells}${conditionalFormatting}${dataValidations}${links.xml}${printSetup}${drawing}${legacyDrawing}${tableParts}</worksheet>`;
	const chunks = streamRowChunks(sheetName, rows, header, footer, rowAttrs, date1904, styles);
	const result: StreamedWorksheet = {
		chunks,
		rels,
		...(comments !== undefined
			? { commentsXml: comments.commentsXml, vmlXml: comments.vmlXml }
			: {}),
		...(pictures !== undefined
			? { drawingXml: pictures.drawingXml, drawingRelsXml: pictures.drawingRelsXml }
			: {}),
		...(tables !== undefined ? { tables } : {}),
		...(autoFilter.ref !== undefined ? { autoFilterRef: autoFilter.ref } : {}),
	};
	Object.setPrototypeOf(result, null);
	return result;
}
