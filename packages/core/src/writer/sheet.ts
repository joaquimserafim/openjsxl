import { XlsxError } from "../errors";
import {
	type CellRef,
	formatRef,
	MAX_COL,
	MAX_COL_WIDTH,
	MAX_ROW,
	MAX_ROW_HEIGHT,
	parseRef,
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
import { MAX_TABLE_NAME_LEN, type TableNameProblem, tableNameProblem } from "../ooxml/table";
import { encodeXstring } from "../ooxml/xstring";
import type {
	ColumnProps,
	Comment,
	ConditionalFormatting,
	DataValidation,
	DataValidationType,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetImage,
	TableInfo,
} from "../types";
import type { MediaRegistry } from "./images";
import { colorXml, type Fail, type StyleRegistry, validateColor } from "./styles";
import type { CellInput, CellValue, SheetInput, StreamSheetInput, StyledCell } from "./types";
import { escapeAttr, escapeText, isPlainRecord, isXmlSafe, preserveAttr } from "./xml";

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

// Split a CellInput into its value, (optional) style, and (optional) formula. Discrimination is
// total: null/undefined are empty, Date instances are dates, primitives are bare values — any OTHER
// object must be a StyledCell. An object with neither a `value` nor a `formula` property is some
// stray object (the pre-F4.2 writer rejected every object; keeping that loudness catches typos like
// { val: 1 } or a nested array). Each caller property is read exactly once (TOCTOU-safe).
function splitInput(
	col: number,
	row: number,
	input: CellInput,
): {
	readonly value: CellValue;
	readonly styled: StyledCell | undefined;
	readonly formula: unknown;
} {
	if (input === null || input === undefined) {
		return { value: input, styled: undefined, formula: undefined };
	}
	if (typeof input !== "object" || input instanceof Date) {
		return { value: input, styled: undefined, formula: undefined };
	}
	const ref = formatRef({ col, row });
	if (Array.isArray(input)) {
		throw new XlsxError("invalid-input", `cell ${ref}: an array is not a cell value`);
	}
	const record = input as unknown as Record<string, unknown>;
	for (const key of Object.keys(record)) {
		if (key !== "value" && key !== "style" && key !== "formula") {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: a cell object allows only "value", "style", and "formula" (got "${key}")`,
			);
		}
	}
	const hasFormula = "formula" in record;
	if (!hasFormula && !("value" in record)) {
		throw new XlsxError(
			"invalid-input",
			`cell ${ref}: an object cell must be { value, style? } or carry a formula`,
		);
	}
	const value = record.value;
	// The inner value must be a plain CellValue — a nested { value } or any other object (except
	// Date) has no meaning and would silently mis-serialize.
	if (
		value !== null &&
		value !== undefined &&
		typeof value === "object" &&
		!(value instanceof Date)
	) {
		throw new XlsxError("invalid-input", `cell ${ref}: a cell's value cannot be an object`);
	}
	return {
		value: value as CellValue,
		styled: input as StyledCell,
		formula: hasFormula ? record.formula : undefined,
	};
}

// Render a single cell, or `undefined` for one that produces no output (an empty value with no
// effective style). Throws `invalid-input` for a value that cannot be represented: a non-finite
// number, an invalid Date, or a type outside the CellInput union (JS callers can pass anything).
// `date1904` selects the serial epoch so it matches the workbook's declared <workbookPr date1904>.
function renderCell(
	col: number,
	row: number,
	input: CellInput,
	date1904: boolean,
	styles: StyleRegistry,
): string | undefined {
	const { value, styled, formula } = splitInput(col, row, input);
	const ref = formatRef({ col, row });

	// Resolve the xf index. Bare non-date values never touch the registry (zero overhead on the
	// unstyled path); a Date always does (it needs the date number format).
	let xf = 0;
	if (value instanceof Date) {
		xf = styles.xfIndexFor(styled?.style, true, ref);
	} else if (styled?.style !== undefined) {
		xf = styles.xfIndexFor(styled.style, false, ref);
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
// XML-safe string in stored form (no leading `=`) within Excel's length ceiling.
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
	return `<c r="${ref}"${sAttr}${cached.tAttr}><f>${escapeText(formula)}</f>${cached.vXml}</c>`;
}

// ── Sheet geometry (F4.5): validation + emission ───────────────────────────────────────────────
// Same philosophy as styles: strict validation naming the sheet, `false`/empty normalize away,
// and the accepted bounds are exactly what the reader's geometry accessors can produce — so the
// bridge's geometry always writes.

function sheetInvalid(sheetName: string, message: string): never {
	throw new XlsxError("invalid-input", `sheet "${sheetName}": ${message}`);
}

function checkKeys(
	sheetName: string,
	what: string,
	obj: Record<string, unknown>,
	allowed: readonly string[],
): void {
	for (const key of Object.keys(obj)) {
		if (!allowed.includes(key))
			sheetInvalid(sheetName, `${what} has an unknown property "${key}"`);
	}
}

// <cols> — one <col> per entry that survives normalization ({hidden: false} alone melts away).
function colsXml(sheetName: string, columns: readonly ColumnProps[] | undefined): string {
	if (columns === undefined) return "";
	if (!Array.isArray(columns)) sheetInvalid(sheetName, "columns must be an array");
	const entries: string[] = [];
	for (let i = 0; i < columns.length; i++) {
		const raw = columns[i] as unknown;
		const what = `columns[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["min", "max", "width", "hidden"]);
		const min = raw.min;
		const max = raw.max;
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
		const width = raw.width;
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
		const hidden = raw.hidden;
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
	for (const key of Object.keys(rowProperties)) {
		const rowNum = Number(key);
		if (!Number.isInteger(rowNum) || rowNum < 1 || rowNum > MAX_ROW) {
			sheetInvalid(
				sheetName,
				`rowProperties key "${key}" is not a row number within Excel's grid`,
			);
		}
		const raw = (rowProperties as Record<string, unknown>)[key];
		const what = `rowProperties[${key}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["height", "hidden"]);
		let attrs = "";
		const height = raw.height;
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
		const hidden = raw.hidden;
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
	const rows = validate(freeze.rows, "rows", MAX_ROW);
	const cols = validate(freeze.cols, "cols", MAX_COL);
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

// Strictly canonical A1: uppercase letters, no leading zeros — the form every real producer
// writes and the only form we emit (parseRef itself would tolerate lowercase). Three letters cap
// the column at ZZZ (18,278), so columnToIndex can't overflow; the grid bound is checked after.
const CANONICAL_CELL = /^[A-Z]{1,3}[1-9][0-9]*$/;

function parseCanonicalRef(ref: string): CellRef | undefined {
	if (!CANONICAL_CELL.test(ref)) return undefined;
	const parsed = parseRef(ref);
	return parsed.col <= MAX_COL && parsed.row <= MAX_ROW ? parsed : undefined;
}

const shortened = (s: string): string => (s.length > 24 ? `${s.slice(0, 24)}…` : s);

interface MergeRect {
	readonly ref: string;
	readonly c1: number;
	readonly r1: number;
	readonly c2: number;
	readonly r2: number;
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
		const ref = merges[i] as unknown;
		const what = `merges[${i}]`;
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what} must be a string`);
		const colon = ref.indexOf(":");
		const from = colon === -1 ? undefined : parseCanonicalRef(ref.slice(0, colon));
		const to = colon === -1 ? undefined : parseCanonicalRef(ref.slice(colon + 1));
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
// and/or an in-workbook location. Empty-string destinations normalize away (the reader never
// yields an empty location); tooltip/display are carried verbatim, empty or not, mirroring the
// reader exactly. Returns the external targets in r:id order — non-empty means the sheet needs
// a relationships part.
function hyperlinksXml(
	sheetName: string,
	hyperlinks: readonly Hyperlink[] | undefined,
): { readonly xml: string; readonly targets: readonly string[] } {
	if (hyperlinks === undefined) return { xml: "", targets: [] };
	if (!Array.isArray(hyperlinks)) sheetInvalid(sheetName, "hyperlinks must be an array");
	const targets: string[] = [];
	const entries: string[] = [];
	for (let i = 0; i < hyperlinks.length; i++) {
		const raw = hyperlinks[i] as unknown;
		const what = `hyperlinks[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["ref", "target", "location", "tooltip", "display"]);
		const ref = raw.ref;
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const colon = ref.indexOf(":");
		const from = parseCanonicalRef(colon === -1 ? ref : ref.slice(0, colon));
		const to = colon === -1 ? from : parseCanonicalRef(ref.slice(colon + 1));
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
		const rawTarget = str("target", raw.target);
		const rawLocation = str("location", raw.location);
		// An empty destination is no destination — the reader drops location="" the same way.
		const target = rawTarget === "" ? undefined : rawTarget;
		const location = rawLocation === "" ? undefined : rawLocation;
		if (target === undefined && location === undefined) {
			sheetInvalid(
				sheetName,
				`${what} needs a target (external) and/or a location (in-workbook)`,
			);
		}
		const tooltip = str("tooltip", raw.tooltip);
		const display = str("display", raw.display);
		let attrs = ` ref="${ref}"`;
		if (target !== undefined) {
			targets.push(target);
			attrs += ` r:id="rId${targets.length}"`;
		}
		if (location !== undefined) attrs += ` location="${escapeAttr(location)}"`;
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
		const raw = dataValidations[i] as unknown;
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
		const sqref = raw.sqref;
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
			const token = sqref[s] as unknown;
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
		const rawType = raw.type;
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

		const formula1 = formulaValue("formula1", raw.formula1);
		const formula2 = formulaValue("formula2", raw.formula2);
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
		const operator = raw.operator;
		if (operator !== undefined) {
			if (!isDataValidationOperator(operator)) {
				sheetInvalid(
					sheetName,
					`${what}.operator must be one of ${DATA_VALIDATION_OPERATORS.join(", ")}`,
				);
			}
			attrs += ` operator="${operator}"`;
		}
		const allowBlank = boolValue("allowBlank", raw.allowBlank);
		if (allowBlank !== undefined) attrs += ` allowBlank="${allowBlank ? 1 : 0}"`;
		const showInputMessage = boolValue("showInputMessage", raw.showInputMessage);
		if (showInputMessage !== undefined)
			attrs += ` showInputMessage="${showInputMessage ? 1 : 0}"`;
		const showErrorMessage = boolValue("showErrorMessage", raw.showErrorMessage);
		if (showErrorMessage !== undefined)
			attrs += ` showErrorMessage="${showErrorMessage ? 1 : 0}"`;
		// showDropDown is INVERTED in the file — intuitive `true` (arrow shown) writes as "0".
		const showDropDown = boolValue("showDropDown", raw.showDropDown);
		if (showDropDown !== undefined) attrs += ` showDropDown="${showDropDown ? 0 : 1}"`;
		const errorStyle = raw.errorStyle;
		if (errorStyle !== undefined) {
			if (!isDataValidationErrorStyle(errorStyle)) {
				sheetInvalid(
					sheetName,
					`${what}.errorStyle must be one of ${DATA_VALIDATION_ERROR_STYLES.join(", ")}`,
				);
			}
			attrs += ` errorStyle="${errorStyle}"`;
		}
		const promptTitle = textValue("promptTitle", raw.promptTitle, MAX_DV_TITLE_LEN);
		if (promptTitle !== undefined) attrs += ` promptTitle="${escapeAttr(promptTitle)}"`;
		const prompt = textValue("prompt", raw.prompt, MAX_DV_TEXT_LEN);
		if (prompt !== undefined) attrs += ` prompt="${escapeAttr(prompt)}"`;
		const errorTitle = textValue("errorTitle", raw.errorTitle, MAX_DV_TITLE_LEN);
		if (errorTitle !== undefined) attrs += ` errorTitle="${escapeAttr(errorTitle)}"`;
		const error = textValue("error", raw.error, MAX_DV_TEXT_LEN);
		if (error !== undefined) attrs += ` error="${escapeAttr(error)}"`;
		// sqref is the LAST attribute (matching Excel/openpyxl); tokens are canonical A1 — no escaping.
		attrs += ` sqref="${tokens.join(" ")}"`;

		let children = "";
		if (formula1 !== undefined) children += `<formula1>${escapeText(formula1)}</formula1>`;
		if (formula2 !== undefined) children += `<formula2>${escapeText(formula2)}</formula2>`;
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

function conditionalFormattingXml(
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
		const rawBlock = cfs[b] as unknown;
		const wb = `conditionalFormatting[${b}]`;
		if (!isPlainRecord(rawBlock)) sheetInvalid(sheetName, `${wb} must be an object`);
		checkKeys(sheetName, wb, rawBlock, ["sqref", "rules"]);

		const sqref = rawBlock.sqref;
		if (!Array.isArray(sqref))
			sheetInvalid(sheetName, `${wb}.sqref must be an array of A1 ranges`);
		if (sqref.length === 0)
			sheetInvalid(sheetName, `${wb}.sqref must cover at least one range`);
		if (sqref.length > MAX_SQREF_RANGES) {
			sheetInvalid(sheetName, `${wb}.sqref has more than ${MAX_SQREF_RANGES} ranges`);
		}
		const tokens: string[] = [];
		for (let s = 0; s < sqref.length; s++) {
			const tok = sqref[s] as unknown;
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

		const rules = rawBlock.rules;
		if (!Array.isArray(rules)) sheetInvalid(sheetName, `${wb}.rules must be an array`);
		if (rules.length === 0) sheetInvalid(sheetName, `${wb}.rules must have at least one rule`);
		for (let r = 0; r < rules.length; r++) {
			built.push(buildCfRule(sheetName, `${wb}.rules[${r}]`, rules[r], styles, b, doc++));
		}
	}

	// Renumber priorities densely 1..n by (caller priority asc, document order asc) — decision 6.
	const order = [...built].sort((a, z) => a.priority - z.priority || a.doc - z.doc);
	const assigned = new Map<number, number>();
	for (let i = 0; i < order.length; i++) {
		const rule = order[i];
		if (rule !== undefined) assigned.set(rule.doc, i + 1);
	}

	let out = "";
	for (let b = 0; b < cfs.length; b++) {
		let rulesXml = "";
		for (const rule of built) {
			if (rule.block !== b) continue;
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

	const type = raw.type;
	if (typeof type !== "string" || !isXmlSafe(type)) {
		sheetInvalid(sheetName, `${what}.type must be an XML-safe string`);
	}
	const priority = raw.priority;
	if (typeof priority !== "number" || !Number.isInteger(priority) || priority < 1) {
		sheetInvalid(sheetName, `${what}.priority must be an integer ≥ 1`);
	}

	// Shared attributes across every variant.
	let rest = "";
	const stopIfTrue = raw.stopIfTrue;
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
		const cfvos = cfvoXml(sheetName, what, raw.cfvo);
		const colors = cfColorsXml(fail, `${what}.colors`, raw.colors);
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
		const cfvos = cfvoXml(sheetName, what, raw.cfvo);
		if (!dataBarCountsOk(cfvos.count)) {
			sheetInvalid(sheetName, `${what} dataBar needs exactly 2 cfvo`);
		}
		const color = colorXml("color", validateColor(fail, `${what}.color`, raw.color));
		return finish('type="dataBar"', "", `<dataBar>${cfvos.xml}${color}</dataBar>`);
	}
	if (type === "iconSet") {
		checkKeys(sheetName, what, raw, ["type", "priority", "stopIfTrue", "iconSet", "cfvo"]);
		let attrs = "";
		const iconSet = raw.iconSet;
		if (iconSet !== undefined) {
			if (typeof iconSet !== "string" || !isXmlSafe(iconSet)) {
				sheetInvalid(sheetName, `${what}.iconSet must be an XML-safe string`);
			}
			attrs += ` iconSet="${escapeAttr(iconSet)}"`;
		}
		const cfvos = cfvoXml(sheetName, what, raw.cfvo);
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
		const dxfId = styles.internDxf(raw.dxf, `${what}.dxf`);
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
		strAttr("operator", raw.operator);
		strAttr("text", raw.text);
		strAttr("timePeriod", raw.timePeriod);
		intAttr("rank", raw.rank);
		boolAttr("percent", raw.percent);
		boolAttr("bottom", raw.bottom);
		boolAttr("aboveAverage", raw.aboveAverage);
		boolAttr("equalAverage", raw.equalAverage);
		intAttr("stdDev", raw.stdDev);
		const children = cfFormulasXml(sheetName, what, raw.formulas);
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
		const c = raw[i] as unknown;
		const cw = `${what}.cfvo[${i}]`;
		if (!isPlainRecord(c)) sheetInvalid(sheetName, `${cw} must be an object`);
		checkKeys(sheetName, cw, c, ["type", "val", "gte"]);
		const t = c.type;
		if (typeof t !== "string" || !isXmlSafe(t)) {
			sheetInvalid(sheetName, `${cw}.type must be an XML-safe string`);
		}
		let attrs = ` type="${escapeAttr(t)}"`;
		const val = c.val;
		if (val !== undefined) {
			if (typeof val !== "string" || !isXmlSafe(val)) {
				sheetInvalid(sheetName, `${cw}.val must be an XML-safe string`);
			}
			attrs += ` val="${escapeAttr(val)}"`;
		}
		const gte = c.gte;
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
		out += colorXml("color", validateColor(fail, `${what}[${i}]`, raw[i]));
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
		const f = raw[i] as unknown;
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
		out += `<formula>${escapeText(stored)}</formula>`;
	}
	return out;
}

// ── Comments (F5.2): xl/commentsN.xml + legacy VML drawing — validation + emission ─────────────
// Excel renders a comment only when the sheet carries BOTH parts: the comments part (an authors
// table + a commentList) AND a legacy VML drawing with one hidden note shape per comment. A
// comments part alone reads back through a tolerant parser but shows nothing in Excel — which is
// why openpyxl always writes both. We emit both from one validated list. Validation mirrors
// hyperlinks: a single-cell canonical `ref` (no ranges), every caller property read exactly once
// (TOCTOU), unknown keys rejected, author/text gated by isXmlSafe. Authors form a
// first-occurrence-ordered unique table; a comment with no author shares one empty-string author
// entry (matching openpyxl). Unused ⇒ nothing emitted, so byte-identity holds for comment-free
// sheets.

// The VML shell is constant; only the note shapes vary. Standard v/o/x prefixes; the box style,
// fill, and shadow are copied verbatim from openpyxl's proven-rendering output.
const VML_HEAD =
	'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
	'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
	'<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
const VML_TAIL = "</xml>";

// One hidden note shape. `index` gives a unique shape id (Excel's ids start at _x0000_s1025) and a
// z-index; `row0`/`col0` are the 0-based cell coordinates. openpyxl emits no explicit <x:Anchor> —
// Excel positions the box from Row/Column plus the fixed margin style — so those coordinates ARE
// the anchor arithmetic.
function vmlShape(index: number, row0: number, col0: number): string {
	return (
		`<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" ` +
		`style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144px;height:79px;z-index:${index + 1};visibility:hidden" ` +
		'fillcolor="#ffffe1" o:insetmode="auto">' +
		'<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/>' +
		'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>' +
		'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill>' +
		`<x:Row>${row0}</x:Row><x:Column>${col0}</x:Column></x:ClientData></v:shape>`
	);
}

interface CommentsParts {
	readonly commentsXml: string;
	readonly vmlXml: string;
}

// Validate the comments and build both parts, or `undefined` when the sheet has none.
function commentsParts(
	sheetName: string,
	comments: readonly Comment[] | undefined,
): CommentsParts | undefined {
	if (comments === undefined) return undefined;
	if (!Array.isArray(comments)) sheetInvalid(sheetName, "comments must be an array");
	if (comments.length === 0) return undefined;
	const authors: string[] = [];
	const authorIndex = new Map<string, number>();
	const items: {
		readonly ref: string;
		readonly authorId: number;
		readonly text: string;
		readonly row: number;
		readonly col: number;
	}[] = [];
	for (let i = 0; i < comments.length; i++) {
		const raw = comments[i] as unknown;
		const what = `comments[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["ref", "author", "text"]);
		// A comment anchors to ONE cell — a range has no meaning (Excel writes a single-cell ref).
		const ref = raw.ref;
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const cell = parseCanonicalRef(ref);
		if (cell === undefined) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" is not a canonical A1 cell within Excel's grid`,
			);
		}
		// text is required (the reader always yields one, possibly empty). Comment text is
		// ST_Xstring content like cell strings: XML-unsafe characters and literal `_xHHHH_`
		// look-alikes store via the escape (F9.6) instead of rejecting — clean text unchanged.
		const rawText = raw.text;
		if (typeof rawText !== "string") sheetInvalid(sheetName, `${what}.text must be a string`);
		const text = encodeXstring(rawText);
		// author is optional; an absent author shares the single empty-string author entry.
		const rawAuthor = raw.author;
		let key: string;
		if (rawAuthor === undefined) key = "";
		else {
			if (typeof rawAuthor !== "string")
				sheetInvalid(sheetName, `${what}.author must be a string`);
			if (!isXmlSafe(rawAuthor)) {
				sheetInvalid(
					sheetName,
					`${what}.author contains a character not allowed in XML (a control character or lone surrogate)`,
				);
			}
			key = rawAuthor;
		}
		let id = authorIndex.get(key);
		if (id === undefined) {
			id = authors.length;
			authors.push(key);
			authorIndex.set(key, id);
		}
		items.push({ ref, authorId: id, text, row: cell.row, col: cell.col });
	}
	const authorsXml = authors.map((a) => `<author>${escapeText(a)}</author>`).join("");
	const listXml = items
		.map(
			(c) =>
				`<comment ref="${c.ref}" authorId="${c.authorId}" shapeId="0"><text><t${preserveAttr(c.text)}>${escapeText(c.text)}</t></text></comment>`,
		)
		.join("");
	const commentsXml = `${XML_DECL}\n<comments xmlns="${NS_MAIN}"><authors>${authorsXml}</authors><commentList>${listXml}</commentList></comments>`;
	const vmlXml = `${VML_HEAD}${items.map((c, j) => vmlShape(j, c.row - 1, c.col - 1)).join("")}${VML_TAIL}`;
	return { commentsXml, vmlXml };
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

// The text value of a header cell: a bare string, or a `{ value: string }` / `{ value: string, style }`
// object. Anything else (number, boolean, formula, blank) is not a valid header — `undefined` here.
function headerCellText(cell: unknown): string | undefined {
	if (typeof cell === "string") return cell;
	if (isPlainRecord(cell) && typeof cell.value === "string") return cell.value;
	return undefined;
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
	const styleName = raw.name;
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
		const value = raw[flag];
		if (value === undefined) continue;
		if (typeof value !== "boolean")
			sheetInvalid(sheetName, `${what}.style.${flag} must be a boolean`);
		attrs += ` ${flag}="${value ? 1 : 0}"`;
	}
	return `<tableStyleInfo${attrs}/>`;
}

/**
 * Validate and emit a sheet's tables (decision 8). Column names DERIVE from the header row when a
 * `headerCell` resolver is supplied (the buffered writer, which has all rows); the streaming writer
 * passes `undefined` and column names come from `tables[i].columns` instead (it can't read the header
 * upfront). `ctx` assigns workbook-global ids and enforces cross-sheet name uniqueness. Returns
 * `undefined` when the sheet has no tables so an unused feature emits nothing (byte-identity).
 */
export function buildTables(
	sheetName: string,
	tables: readonly TableInfo[] | undefined,
	headerCell: ((row: number, col: number) => unknown) | undefined,
	ctx: TableContext,
	styles: StyleRegistry,
): TablePart[] | undefined {
	if (tables === undefined) return undefined;
	if (!Array.isArray(tables)) sheetInvalid(sheetName, "tables must be an array");
	if (tables.length === 0) return undefined;
	const rects: { c1: number; r1: number; c2: number; r2: number; name: string }[] = [];
	const parts: TablePart[] = [];

	for (let i = 0; i < tables.length; i++) {
		const raw = tables[i] as unknown;
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

		const name = raw.name;
		if (typeof name !== "string") sheetInvalid(sheetName, `${what}.name must be a string`);
		validateTableName(sheetName, what, name);
		ctx.claimName(sheetName, name);

		const ref = raw.ref;
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const colon = ref.indexOf(":");
		const from = colon === -1 ? undefined : parseCanonicalRef(ref.slice(0, colon));
		const to = colon === -1 ? undefined : parseCanonicalRef(ref.slice(colon + 1));
		if (from === undefined || to === undefined || to.col < from.col || to.row < from.row) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" is not a canonical A1 range like "A1:C5" within Excel's grid`,
			);
		}
		const width = to.col - from.col + 1;

		const headerRow = raw.headerRow !== false; // default true
		const totalsRow = raw.totalsRow === true;

		const columns = raw.columns;
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
			// Read the caller's column entry ONCE (single-read TOCTOU — a Proxy array must not return a
			// different object between the isPlainRecord check and its use).
			const rawColumn = columns !== undefined ? columns[c] : undefined;
			const provided = isPlainRecord(rawColumn) ? rawColumn : undefined;
			let columnName: string;
			if (headerRow && headerCell !== undefined) {
				const text = headerCellText(headerCell(from.row, from.col + c));
				const provName = provided !== undefined ? provided.name : undefined;
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
				const provName = provided !== undefined ? provided.name : undefined;
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
		const styleXml = tableStyleInfoXml(sheetName, what, raw.style);
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
	const header = styles.internDxf(raw.headerRowStyle, `${what}.headerRowStyle`);
	if (header !== undefined) attrs += ` headerRowDxfId="${header}"`;
	const data = styles.internDxf(raw.dataStyle, `${what}.dataStyle`);
	if (data !== undefined) attrs += ` dataDxfId="${data}"`;
	const totals = styles.internDxf(raw.totalsRowStyle, `${what}.totalsRowStyle`);
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
		const label = provided.totalsRowLabel;
		if (label !== undefined) {
			if (typeof label !== "string")
				sheetInvalid(sheetName, `${what} totalsRowLabel must be a string`);
			attrs += ` totalsRowLabel="${escapeAttr(encodeXstring(label))}"`;
		}
		const fn = provided.totalsRowFunction;
		if (fn !== undefined)
			attrs += ` totalsRowFunction="${escapeAttr(requireXmlString(sheetName, `${what} totalsRowFunction`, fn))}"`;
		const totalsFormula = provided.totalsRowFormula;
		if (totalsFormula !== undefined) {
			children += `<totalsRowFormula>${escapeText(requireXmlString(sheetName, `${what} totalsRowFormula`, totalsFormula))}</totalsRowFormula>`;
		}
		const calcFormula = provided.calculatedColumnFormula;
		if (calcFormula !== undefined) {
			children += `<calculatedColumnFormula>${escapeText(requireXmlString(sheetName, `${what} calculatedColumnFormula`, calcFormula))}</calculatedColumnFormula>`;
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
	const col = raw.col;
	const row = raw.row;
	if (typeof col !== "number" || !Number.isInteger(col) || col < 1 || col > MAX_COL) {
		sheetInvalid(sheetName, `${what}.col must be an integer column in 1..${MAX_COL}`);
	}
	if (typeof row !== "number" || !Number.isInteger(row) || row < 1 || row > MAX_ROW) {
		sheetInvalid(sheetName, `${what}.row must be an integer row in 1..${MAX_ROW}`);
	}
	return {
		col,
		row,
		colOff: emuValue(sheetName, `${what}.colOff`, raw.colOff),
		rowOff: emuValue(sheetName, `${what}.rowOff`, raw.rowOff),
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
		const raw = images[i] as unknown;
		const what = `images[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["anchor", "bytes", "mime", "name"]);

		// bytes — read the reference ONCE (a getter must not swap the buffer between here and packing).
		const bytes = raw.bytes;
		if (!(bytes instanceof Uint8Array) || bytes.length === 0) {
			sheetInvalid(sheetName, `${what}.bytes must be a non-empty Uint8Array`);
		}
		// mime — allowlist only; the extension is derived from the mime, never sniffed from the data.
		// Use hasOwn, not a bare lookup: a bare `map[mime]` walks the prototype chain, so a mime like
		// "constructor" or "toString" would resolve to an inherited value and slip past the gate.
		const mime = raw.mime;
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
		const rawName = raw.name;
		if (rawName !== undefined && (typeof rawName !== "string" || !isXmlSafe(rawName))) {
			sheetInvalid(sheetName, `${what}.name must be an XML-safe string`);
		}
		const name = rawName ?? `Image ${i + 1}`;

		// anchor — exactly one of `to` (two-cell) or `ext` (one-cell).
		const anchor = raw.anchor;
		if (!isPlainRecord(anchor)) sheetInvalid(sheetName, `${what}.anchor must be an object`);
		checkKeys(sheetName, `${what}.anchor`, anchor, ["from", "to", "ext", "editAs"]);
		const from = validatePoint(sheetName, `${what}.anchor.from`, anchor.from);
		// Read `to`/`ext` ONCE each (single-read TOCTOU): a getter must not present one shape to the
		// XOR check and another to emission.
		const rawTo = anchor.to;
		const rawExt = anchor.ext;
		const hasTo = rawTo !== undefined;
		const hasExt = rawExt !== undefined;
		if (hasTo === hasExt) {
			sheetInvalid(
				sheetName,
				`${what}.anchor must have exactly one of "to" (two-cell) or "ext" (one-cell)`,
			);
		}
		const editAs = anchor.editAs;
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
			const cx = emuValue(sheetName, `${what}.anchor.ext.cx`, rawExt.cx);
			const cy = emuValue(sheetName, `${what}.anchor.ext.cy`, rawExt.cy);
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
 * `sheetIndex` (0-based) names the sheet's own extra parts (comments, VML) so their rel targets and
 * package paths agree.
 */
export function worksheetXml(
	sheet: SheetInput,
	sheetIndex: number,
	date1904: boolean,
	styles: StyleRegistry,
	media: MediaRegistry,
	tableCtx: TableContext,
): WorksheetResult {
	const rows = sheet.rows;
	// Geometry and metadata validate up front (geometry also contributes rows below): a bad
	// column/row/freeze/merge/hyperlink/comment/image spec must surface before any cell work.
	const cols = colsXml(sheet.name, sheet.columns);
	const rowAttrs = rowAttrsMap(sheet.name, sheet.rowProperties);
	const sheetViews = sheetViewsXml(sheet.name, sheet.freeze);
	const mergeCells = mergeCellsXml(sheet.name, sheet.merges);
	const conditionalFormatting = conditionalFormattingXml(
		sheet.name,
		sheet.conditionalFormatting,
		styles,
	);
	const dataValidations = dataValidationsXml(sheet.name, sheet.dataValidations);
	const links = hyperlinksXml(sheet.name, sheet.hyperlinks);
	const comments = commentsParts(sheet.name, sheet.comments);
	const pictures = imageParts(sheet.name, sheet.images, media);
	// Column names derive from the header row — the buffered writer has every row, so it reads them.
	const tables = buildTables(
		sheet.name,
		sheet.tables,
		(row, col) => {
			const rowCells = rows[row - 1];
			return Array.isArray(rowCells) ? rowCells[col - 1] : undefined;
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

	for (let r = 0; r < rows.length; r++) {
		const cells = rows[r];
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
		if (cells.length === 0) continue;
		if (cells.length > MAX_COL) {
			throw new XlsxError(
				"invalid-input",
				`sheet row ${r + 1}: a row cannot have more than ${MAX_COL} cells`,
			);
		}
		const rowNum = r + 1;
		const cellXmls: string[] = [];
		for (let c = 0; c < cells.length; c++) {
			const colNum = c + 1;
			const rendered = renderCell(colNum, rowNum, cells[c], date1904, styles);
			if (rendered === undefined) continue;
			if (minRow === 0 || rowNum < minRow) minRow = rowNum;
			if (rowNum > maxRow) maxRow = rowNum;
			if (minCol === 0 || colNum < minCol) minCol = colNum;
			if (colNum > maxCol) maxCol = colNum;
			cellXmls.push(rendered);
		}
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

	// Schema order within <worksheet>: dimension, sheetViews, cols, sheetData, mergeCells,
	// conditionalFormatting, dataValidations, hyperlinks, drawing, legacyDrawing, tableParts
	// (CT_Worksheet sequence — conditionalFormatting then dataValidations sit between mergeCells and
	// hyperlinks; tableParts is last). Every optional block is an empty string when unused, so a sheet
	// using none of them emits the exact pre-F4.5/F4.6/F5.2/F9.1/F9.2/F9.3 bytes.
	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"${nsR}><dimension ref="${dimension}"/>${sheetViews}${cols}<sheetData>${rowXmls
		.map(([, x]) => x)
		.join(
			"",
		)}</sheetData>${mergeCells}${conditionalFormatting}${dataValidations}${links.xml}${drawing}${legacyDrawing}${tableParts}</worksheet>`;
	// Comment/drawing/table side parts are spread only when present so the optional properties stay
	// truly absent (exactOptionalPropertyTypes).
	return {
		xml,
		rels,
		...(comments !== undefined
			? { commentsXml: comments.commentsXml, vmlXml: comments.vmlXml }
			: {}),
		...(pictures !== undefined
			? { drawingXml: pictures.drawingXml, drawingRelsXml: pictures.drawingRelsXml }
			: {}),
		...(tables !== undefined ? { tables } : {}),
	};
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

// Render one streamed row. Its number is the 1-based position in the stream. Mirrors the buffered
// per-row logic: an empty row with no properties emits nothing; a property-only row emits `<row r/>`;
// otherwise `<row r>cells</row>`. Interns styles into the shared registry as it renders.
function renderStreamRow(
	sheetName: string,
	rowNum: number,
	cells: readonly CellInput[] | undefined,
	rowAttrs: Map<number, string>,
	date1904: boolean,
	styles: StyleRegistry,
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
	if (cells.length > MAX_COL) {
		throw new XlsxError(
			"invalid-input",
			`sheet "${sheetName}" row ${rowNum}: a row cannot have more than ${MAX_COL} cells`,
		);
	}
	const cellXmls: string[] = [];
	for (let c = 0; c < cells.length; c++) {
		const rendered = renderCell(c + 1, rowNum, cells[c], date1904, styles);
		if (rendered !== undefined) cellXmls.push(rendered);
	}
	if (cellXmls.length === 0) return attrs !== "" ? `<row r="${rowNum}"${attrs}/>` : "";
	return `<row r="${rowNum}"${attrs}>${cellXmls.join("")}</row>`;
}

// Roughly this many bytes of row XML accumulate before a chunk is flushed — batching keeps the
// compressor fed with substantial writes instead of one tiny write per row.
const ROW_CHUNK_BYTES = 65536;

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
	let buf = "";
	// `for await` iterates both a sync iterable (an array) and an async one (a DB cursor).
	for await (const cells of rows) {
		rowNum++;
		if (rowNum > MAX_ROW) {
			throw new XlsxError(
				"invalid-input",
				`sheet "${sheetName}": a sheet cannot have more than ${MAX_ROW} rows`,
			);
		}
		buf += renderStreamRow(sheetName, rowNum, cells, rowAttrs, date1904, styles);
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
 */
export function streamWorksheet(
	sheet: StreamSheetInput,
	sheetIndex: number,
	date1904: boolean,
	styles: StyleRegistry,
	media: MediaRegistry,
	tableCtx: TableContext,
): StreamedWorksheet {
	const cols = colsXml(sheet.name, sheet.columns);
	const rowAttrs = rowAttrsMap(sheet.name, sheet.rowProperties);
	const sheetViews = sheetViewsXml(sheet.name, sheet.freeze);
	const mergeCells = mergeCellsXml(sheet.name, sheet.merges);
	const conditionalFormatting = conditionalFormattingXml(
		sheet.name,
		sheet.conditionalFormatting,
		styles,
	);
	const dataValidations = dataValidationsXml(sheet.name, sheet.dataValidations);
	const links = hyperlinksXml(sheet.name, sheet.hyperlinks);
	const comments = commentsParts(sheet.name, sheet.comments);
	const pictures = imageParts(sheet.name, sheet.images, media);
	// The footer (which carries <tableParts>) is built BEFORE the rows stream, so the header row isn't
	// available — column names come from `tables[i].columns` here (no header resolver).
	const tables = buildTables(sheet.name, sheet.tables, undefined, tableCtx, styles);

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
	const footer = `</sheetData>${mergeCells}${conditionalFormatting}${dataValidations}${links.xml}${drawing}${legacyDrawing}${tableParts}</worksheet>`;
	const chunks = streamRowChunks(
		sheet.name,
		sheet.rows,
		header,
		footer,
		rowAttrs,
		date1904,
		styles,
	);
	return {
		chunks,
		rels,
		...(comments !== undefined
			? { commentsXml: comments.commentsXml, vmlXml: comments.vmlXml }
			: {}),
		...(pictures !== undefined
			? { drawingXml: pictures.drawingXml, drawingRelsXml: pictures.drawingRelsXml }
			: {}),
		...(tables !== undefined ? { tables } : {}),
	};
}
