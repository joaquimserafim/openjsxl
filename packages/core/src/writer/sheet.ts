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
import { dateToSerial } from "../ooxml/dates";
import { MAX_EMU, MEDIA_MIME_TO_EXT } from "../ooxml/drawing";
import { MAX_FORMULA_LEN } from "../ooxml/formula";
import type { ColumnProps, Comment, FreezePane, Hyperlink, RowProps, SheetImage } from "../types";
import type { MediaRegistry } from "./images";
import type { StyleRegistry } from "./styles";
import type { CellInput, CellValue, SheetInput, StreamSheetInput, StyledCell } from "./types";
import { escapeAttr, escapeText, isXmlSafe, preserveAttr } from "./xml";

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
		// A forbidden control character or lone surrogate would make the part not well-formed (or be
		// silently mangled to U+FFFD by TextEncoder) — reject rather than emit a broken/lossy file.
		if (!isXmlSafe(value)) {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: string contains a character not allowed in XML (a control character or lone surrogate)`,
			);
		}
		return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${preserveAttr(value)}>${escapeText(value)}</t></is></c>`;
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
		if (!isXmlSafe(value)) {
			throw new XlsxError(
				"invalid-input",
				`cell ${ref}: cached string contains a character not allowed in XML (a control character or lone surrogate)`,
			);
		}
		return { tAttr: ' t="str"', vXml: `<v>${escapeText(value)}</v>` };
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

function isPlainRecord(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false;
	const proto = Object.getPrototypeOf(value);
	return proto === null || proto === Object.prototype;
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
		// text is required (the reader always yields one, possibly empty) and must be XML-safe.
		const text = raw.text;
		if (typeof text !== "string") sheetInvalid(sheetName, `${what}.text must be a string`);
		if (!isXmlSafe(text)) {
			sheetInvalid(
				sheetName,
				`${what}.text contains a character not allowed in XML (a control character or lone surrogate)`,
			);
		}
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
): { rels: SheetRel[]; drawing: string; legacyDrawing: string; nsR: string } {
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
	// xmlns:r is declared whenever the body references an rId — a hyperlink, drawing, or legacyDrawing.
	// A sheet using none keeps the exact pre-F4.6 root element.
	const nsR =
		hyperlinkTargets.length > 0 || drawing !== "" || legacyDrawing !== ""
			? ` xmlns:r="${NS_REL}"`
			: "";
	return { rels, drawing, legacyDrawing, nsR };
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
): WorksheetResult {
	const rows = sheet.rows;
	// Geometry and metadata validate up front (geometry also contributes rows below): a bad
	// column/row/freeze/merge/hyperlink/comment/image spec must surface before any cell work.
	const cols = colsXml(sheet.name, sheet.columns);
	const rowAttrs = rowAttrsMap(sheet.name, sheet.rowProperties);
	const sheetViews = sheetViewsXml(sheet.name, sheet.freeze);
	const mergeCells = mergeCellsXml(sheet.name, sheet.merges);
	const links = hyperlinksXml(sheet.name, sheet.hyperlinks);
	const comments = commentsParts(sheet.name, sheet.comments);
	const pictures = imageParts(sheet.name, sheet.images, media);

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
	const { rels, drawing, legacyDrawing, nsR } = sheetRelPlumbing(
		sheetIndex,
		links.targets,
		comments !== undefined,
		pictures !== undefined,
	);

	// Schema order within <worksheet>: dimension, sheetViews, cols, sheetData, mergeCells,
	// hyperlinks, drawing, legacyDrawing (CT_Worksheet sequence — drawing precedes legacyDrawing).
	// Every optional block is an empty string when unused, so a sheet using none of them emits the
	// exact pre-F4.5/F4.6/F5.2 bytes.
	const xml = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"${nsR}><dimension ref="${dimension}"/>${sheetViews}${cols}<sheetData>${rowXmls
		.map(([, x]) => x)
		.join("")}</sheetData>${mergeCells}${links.xml}${drawing}${legacyDrawing}</worksheet>`;
	// Comment/drawing side parts are spread only when present so the optional properties stay truly
	// absent (exactOptionalPropertyTypes).
	return {
		xml,
		rels,
		...(comments !== undefined
			? { commentsXml: comments.commentsXml, vmlXml: comments.vmlXml }
			: {}),
		...(pictures !== undefined
			? { drawingXml: pictures.drawingXml, drawingRelsXml: pictures.drawingRelsXml }
			: {}),
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
): StreamedWorksheet {
	const cols = colsXml(sheet.name, sheet.columns);
	const rowAttrs = rowAttrsMap(sheet.name, sheet.rowProperties);
	const sheetViews = sheetViewsXml(sheet.name, sheet.freeze);
	const mergeCells = mergeCellsXml(sheet.name, sheet.merges);
	const links = hyperlinksXml(sheet.name, sheet.hyperlinks);
	const comments = commentsParts(sheet.name, sheet.comments);
	const pictures = imageParts(sheet.name, sheet.images, media);

	// Relationships + the body plumbing that references them — the SAME builder the buffered writer
	// uses, so the two can't drift on rel order, rId allocation, or the drawing/legacyDrawing refs.
	const { rels, drawing, legacyDrawing, nsR } = sheetRelPlumbing(
		sheetIndex,
		links.targets,
		comments !== undefined,
		pictures !== undefined,
	);

	const header = `${XML_DECL}\n<worksheet xmlns="${NS_MAIN}"${nsR}>${sheetViews}${cols}<sheetData>`;
	const footer = `</sheetData>${mergeCells}${links.xml}${drawing}${legacyDrawing}</worksheet>`;
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
	};
}
