import type {
	Alignment,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	CellProtection,
	CellStyle,
	Color,
	FillStyle,
	FontStyle,
	HorizontalAlignment,
	PatternType,
	UnderlineStyle,
	VerticalAlignment,
} from "../types";
import { isXmlSafe, localName } from "../utils";
import { tokenize } from "../xml";

// xl/styles.xml — the workbook's style tables. A cell's `s` attribute indexes `<cellXfs>`; each
// `<xf>` names a numFmtId plus fontId/fillId/borderId into the component tables and may carry an
// inline `<alignment>`. Two consumers read this:
//
//   - The HOT value path (F2.1): `isDateStyle`/`formatCode` — numFmtId only, decides whether a
//     number is a date. Untouched by the style model below.
//   - The LAZY style path (F4.1): `cellStyle(i)` materializes the full resolved CellStyle for an
//     xf index on first request and caches it — one reference-stable object per distinct xf, so
//     the writer's interner (and the bridge) can dedup by identity.
//
//   numFmtId < 164  built-in, NOT written into <numFmts>. The date/time ones are fixed by
//                   the spec (14–22, 45–47) plus the locale date/time block (27–36, 50–58).
//   numFmtId ≥ 164  custom, defined in <numFmts> with a formatCode we sniff for date tokens.
//
// Deliberate simplifications (openpyxl-compatible, documented in IMPLEMENTATION.md M4): the
// `apply*` flags and `cellStyleXfs`/`xfId` named-style inheritance are ignored — the component
// ids on the cellXf itself are taken as effective, which is what every mainstream producer
// writes anyway. Gradient fills, diagonal borders, accounting underlines, and the legacy
// textRotation=255 marker degrade to "absent" rather than inventing a lossy approximation.

export interface StyleTable {
	/** True when the cell format at this `s` index applies a date/time number format. */
	isDateStyle(styleIndex: number | undefined): boolean;
	/**
	 * The number-format code applied at this `s` index — a custom code (`<numFmts>`) or a
	 * built-in one (e.g. `"0.00%"`, `"mm-dd-yy"`). `undefined` when the id is a locale/reserved
	 * built-in with no portable code, or the index is out of range.
	 */
	formatCode(styleIndex: number | undefined): string | undefined;
	/**
	 * The full resolved style at this `s` index — number format code, font, fill, border, and
	 * alignment. `undefined` when the index is out of range or the xf resolves to the workbook
	 * default (no distinguishing component). Cached: repeated calls return the same object.
	 */
	cellStyle(styleIndex: number | undefined): CellStyle | undefined;
}

// Built-in number formats (ECMA-376 §18.8.30) with a fixed, non-locale code. The locale and
// reserved ids (23–36, 41–44, 50–58) have no portable code and resolve to undefined; date
// detection still recognises those via isBuiltinDateId, we just don't fabricate their string.
// Exported: the writer reverse-maps exact code matches to these ids (F4.3), so the two tables
// can never drift.
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
	0: "General",
	1: "0",
	2: "0.00",
	3: "#,##0",
	4: "#,##0.00",
	5: '"$"#,##0_);("$"#,##0)',
	6: '"$"#,##0_);[Red]("$"#,##0)',
	7: '"$"#,##0.00_);("$"#,##0.00)',
	8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
	9: "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0_);(#,##0)",
	38: "#,##0_);[Red](#,##0)",
	39: "#,##0.00_);(#,##0.00)",
	40: "#,##0.00_);[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0E+0",
	49: "@",
};

export function isBuiltinDateId(id: number): boolean {
	return (
		(id >= 14 && id <= 22) ||
		(id >= 27 && id <= 36) ||
		(id >= 45 && id <= 47) ||
		(id >= 50 && id <= 58)
	);
}

// A format code is a date/time format when, after removing the parts that can never be date
// tokens — quoted literals ("…"), escaped characters (\x), skip/fill tokens (_x, *x), and
// bracketed sections ([Red], [$-409], [>100]) — one of the date/time letters d m y h s remains.
//
// Elapsed-time tokens — [h], [hh], [mm], [ss] — are an exception: they ARE time formats but
// live inside brackets the bracket pass would strip, so a code like "[h]" (or "[mm]:[ss]")
// would otherwise reduce to nothing. Detect them AFTER stripping literals but BEFORE stripping
// brackets: testing the raw code instead would let a quoted literal containing "[h]" (e.g.
// `"[h] rate" 0.00`, a plain number format) masquerade as elapsed time — a written cell would
// then re-read as a date while Excel shows a number (adversarial review, F4.3).
const ELAPSED_TIME = /\[(?:h+|m+|s+)\]/i;
// Quoted literals, backslash escapes, and the ECMA-376 skip (_x) / fill (*x) tokens: their
// characters render literally (or pad), so an m/d/h inside them is not a date token ("0.00_m").
const LITERALS = /"[^"]*"|\\.|[_*]./g;
const BRACKETS = /\[[^\]]*\]/g;
const DATE_TOKEN = /[dmyhs]/i;

export function isDateFormatCode(formatCode: string): boolean {
	const withoutLiterals = formatCode.replace(LITERALS, "");
	if (ELAPSED_TIME.test(withoutLiterals)) return true;
	return DATE_TOKEN.test(withoutLiterals.replace(BRACKETS, ""));
}

// ── Attribute parsing helpers ──────────────────────────────────────────────────────────────────

// OOXML boolean attributes: an element like <b/> asserts true by presence; an explicit
// val="0"/"false" negates it. Anything else (including absence of the element) is handled by the
// caller — this decides only what a PRESENT element with this `val` means. Exported so the dxf
// parser (F9.3) shares the exact same semantics.
export function boolAttr(val: string | undefined): boolean {
	return val !== "0" && val !== "false";
}

// Shared read/write bounds for style values. The reader DEGRADES anything outside them (a
// tolerated file keeps opening; the malformed piece reads as absent) and the writer REJECTS them
// (invalid-input) — sharing the constants makes the two ends of the bridge agree by construction:
// whatever cellStyle() emits, the writer's validators accept, so read → modify → write can never
// crash on a style the reader was happy to produce (adversarial review, F4.4).
/** 6-digit RGB or 8-digit ARGB hex — the only rgb forms real files carry. */
export const HEX_COLOR = /^(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/;
/** xsd:unsignedInt ceiling for theme/indexed color values. */
export const MAX_COLOR_INDEX = 0xffffffff;
/** Excel's own maximum alignment indent. */
export const MAX_INDENT = 250;
/**
 * xsd:unsignedInt ceiling for a password-hash `spinCount`/`workbookSpinCount` (F10.3). SHARED bound:
 * the tolerant reader DROPS a spinCount above it (a value Excel could never have written), and the
 * strict writer REJECTS one typed — so a hostile 21-digit count can't slip through as `1e+21` (invalid
 * integer XML that would make Excel repair-prompt) or as a silently precision-lost integer.
 */
export const MAX_SPIN_COUNT = 0xffffffff;
/**
 * Page-margin ceiling in inches (F10.4). Excel's own maximum is 49"; larger values (or non-finite ones)
 * are hostile. SHARED bound: the reader clamps a finite margin into `[0, MAX_PAGE_MARGIN]` (and drops the
 * whole `<pageMargins>` if a value is non-numeric/non-finite), the writer rejects one outside the range.
 */
export const MAX_PAGE_MARGIN = 49;
/** Print-scale bounds as a percentage (F10.4, `<pageSetup scale>`): 10–400. Reader clamps, writer rejects. */
export const MIN_PAGE_SCALE = 10;
export const MAX_PAGE_SCALE = 400;

// A <color> / <fgColor> / <bgColor> element's attributes → the raw Color union. The spec allows
// exactly one addressing mode per element; if a producer emits several, precedence follows what
// consumers (and openpyxl) do: rgb > theme > indexed > auto. Malformed values — bad numerics,
// non-hex rgb, out-of-range indexes — yield undefined (no color) rather than an unwritable record.
export function parseColor(attrs: Readonly<Record<string, string | undefined>>): Color | undefined {
	if (attrs.rgb !== undefined) {
		return HEX_COLOR.test(attrs.rgb) ? { rgb: attrs.rgb } : undefined;
	}
	if (attrs.theme !== undefined) {
		const theme = Number(attrs.theme);
		if (!Number.isInteger(theme) || theme < 0 || theme > MAX_COLOR_INDEX) return undefined;
		const tint = attrs.tint === undefined ? undefined : Number(attrs.tint);
		if (tint !== undefined && Number.isFinite(tint)) return { theme, tint };
		return { theme };
	}
	if (attrs.indexed !== undefined) {
		const indexed = Number(attrs.indexed);
		return Number.isInteger(indexed) && indexed >= 0 && indexed <= MAX_COLOR_INDEX
			? { indexed }
			: undefined;
	}
	if (attrs.auto !== undefined && boolAttr(attrs.auto)) return { auto: true };
	return undefined;
}

// Enum whitelists: attribute values are producer-controlled strings, so gate them through the
// spec's literal sets — garbage degrades to "absent", never leaks into the typed model. Exported:
// the writer validates its style input against the SAME sets, so reader and writer cannot drift.
export const PATTERN_TYPES = new Set<PatternType>([
	"none",
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625",
]);
export const BORDER_LINE_STYLES = new Set<BorderLineStyle>([
	"thin",
	"medium",
	"thick",
	"dashed",
	"dotted",
	"double",
	"hair",
	"mediumDashed",
	"dashDot",
	"mediumDashDot",
	"dashDotDot",
	"mediumDashDotDot",
	"slantDashDot",
]);
export const H_ALIGNMENTS = new Set<HorizontalAlignment>([
	"left",
	"center",
	"right",
	"justify",
	"fill",
	"centerContinuous",
	"distributed",
]);
export const V_ALIGNMENTS = new Set<VerticalAlignment>([
	"top",
	"center",
	"bottom",
	"justify",
	"distributed",
]);

// The <font> child elements the model reads. The font-children dispatch branch is gated on THIS
// set (not just "a font is open") so a dangling unclosed <font> can never swallow structural
// tokens like <fills> or <xf> — see the dispatch chain below.
const FONT_CHILDREN = new Set(["name", "sz", "b", "i", "u", "strike", "color"]);

export function parseAlignment(
	attrs: Readonly<Record<string, string | undefined>>,
): Alignment | undefined {
	const out: {
		horizontal?: HorizontalAlignment;
		vertical?: VerticalAlignment;
		wrapText?: boolean;
		shrinkToFit?: boolean;
		indent?: number;
		textRotation?: number;
	} = {};
	if (
		attrs.horizontal !== undefined &&
		H_ALIGNMENTS.has(attrs.horizontal as HorizontalAlignment)
	) {
		out.horizontal = attrs.horizontal as HorizontalAlignment;
	}
	if (attrs.vertical !== undefined && V_ALIGNMENTS.has(attrs.vertical as VerticalAlignment)) {
		out.vertical = attrs.vertical as VerticalAlignment;
	}
	if (attrs.wrapText !== undefined && boolAttr(attrs.wrapText)) out.wrapText = true;
	if (attrs.shrinkToFit !== undefined && boolAttr(attrs.shrinkToFit)) out.shrinkToFit = true;
	if (attrs.indent !== undefined) {
		const indent = Number(attrs.indent);
		// Excel's own indent ceiling; a larger value in a file is producer garbage — degrade.
		if (Number.isInteger(indent) && indent > 0 && indent <= MAX_INDENT) out.indent = indent;
	}
	if (attrs.textRotation !== undefined) {
		// 0–180 per the spec (91–180 = downward). The legacy 255 "vertical stacked" marker is not
		// modelled — degrade to no rotation rather than misrepresent it as 255 degrees.
		const rotation = Number(attrs.textRotation);
		if (Number.isInteger(rotation) && rotation > 0 && rotation <= 180)
			out.textRotation = rotation;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

const hasKeys = (o: object): boolean => Object.keys(o).length > 0;

// One cellXfs <xf> as read: the numFmt id the hot path needs plus the component table ids and
// inline alignment the style path resolves through.
interface XfRecord {
	readonly numFmtId: number;
	readonly fontId: number;
	readonly fillId: number;
	readonly borderId: number;
	readonly alignment: Alignment | undefined;
	readonly protection: CellProtection | undefined;
}

// A cellXfs <xf>'s inline <protection locked hidden> (F10.3), or undefined when it has neither flag.
// Both flags are read only when explicitly present, so what the reader returns re-writes byte-faithfully.
function parseProtection(attrs: Readonly<Record<string, string>>): CellProtection | undefined {
	const out: { locked?: boolean; hidden?: boolean } = {};
	if (attrs.locked === "1" || attrs.locked === "true") out.locked = true;
	else if (attrs.locked === "0" || attrs.locked === "false") out.locked = false;
	if (attrs.hidden === "1" || attrs.hidden === "true") out.hidden = true;
	else if (attrs.hidden === "0" || attrs.hidden === "false") out.hidden = false;
	return out.locked !== undefined || out.hidden !== undefined ? out : undefined;
}

export function parseStyles(xml: string): StyleTable {
	const customFormats = new Map<number, string>();
	const xfs: XfRecord[] = [];
	// Component tables by index. A slot is undefined when the record isn't representable in the
	// model (a gradient fill) — a cell referencing it simply has no such component.
	const fonts: FontStyle[] = [];
	const fills: (FillStyle | undefined)[] = [];
	const borders: BorderStyle[] = [];

	let inNumFmts = false;
	let inCellXfs = false; // NOT cellStyleXfs — a cell's `s` indexes cellXfs only
	let inXf = false; // inside a non-self-closing cellXfs <xf>, awaiting <alignment>

	// <fonts> builder state: the font record under construction.
	let inFonts = false;
	let font:
		| {
				name?: string;
				size?: number;
				bold?: boolean;
				italic?: boolean;
				underline?: UnderlineStyle;
				strike?: boolean;
				color?: Color;
		  }
		| undefined;

	// <fills> builder state. `fill` is set by <patternFill> (and stays undefined for
	// <gradientFill>, so the slot records "not representable").
	let inFills = false;
	let inFill = false;
	let fill: { patternType: PatternType; fgColor?: Color; bgColor?: Color } | undefined;

	// <borders> builder state. `edge` holds a started edge element awaiting an optional <color>
	// child; edges without a style attribute are not borders and never commit.
	let inBorders = false;
	let border:
		| { top?: BorderEdge; right?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge }
		| undefined;
	let edgeName: "top" | "right" | "bottom" | "left" | undefined;
	let edgeStyle: BorderLineStyle | undefined;
	let edgeColor: Color | undefined;

	const commitEdge = (): void => {
		if (border !== undefined && edgeName !== undefined && edgeStyle !== undefined) {
			border[edgeName] =
				edgeColor !== undefined
					? { style: edgeStyle, color: edgeColor }
					: { style: edgeStyle };
		}
		edgeName = undefined;
		edgeStyle = undefined;
		edgeColor = undefined;
	};

	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		const name = localName(token.name);

		if (token.kind === "open") {
			if (name === "numFmts") {
				if (!token.selfClosing) inNumFmts = true;
			} else if (name === "numFmt" && inNumFmts) {
				const id = Number(token.attrs.numFmtId);
				const code = token.attrs.formatCode;
				if (Number.isInteger(id) && code !== undefined) customFormats.set(id, code);
			} else if (name === "fonts") {
				if (!token.selfClosing) inFonts = true;
			} else if (name === "font" && inFonts) {
				font = {};
				if (token.selfClosing) {
					fonts.push({});
					font = undefined;
				}
			} else if (font !== undefined && FONT_CHILDREN.has(name)) {
				// Children of the open <font>, gated by NAME so this branch can never swallow a
				// structural token. Without the gate, a dangling unclosed <font> (misnested input
				// the non-validating tokenizer passes through) would intercept <fills>, <cellXfs>,
				// every <xf>, … — silently emptying the tables and regressing even the date hot
				// path (adversarial review, F4.1). Unmodelled font children (family, scheme,
				// charset, vertAlign) simply aren't in the set and fall through harmlessly.
				if (name === "name" && token.attrs.val !== undefined) {
					// An empty or XML-unsafe name (a control character can arrive via a decoded
					// numeric reference) is unwritable — degrade to "no per-cell font name".
					const val = token.attrs.val;
					if (val !== "" && isXmlSafe(val)) font.name = val;
				} else if (name === "sz") {
					const size = Number(token.attrs.val);
					if (Number.isFinite(size) && size > 0) font.size = size;
				} else if (name === "b") {
					if (boolAttr(token.attrs.val)) font.bold = true;
				} else if (name === "i") {
					if (boolAttr(token.attrs.val)) font.italic = true;
				} else if (name === "u") {
					// <u/> is single; val names the variant. Accounting variants degrade (deferred).
					const val = token.attrs.val ?? "single";
					if (val === "single" || val === "double") font.underline = val;
				} else if (name === "strike") {
					if (boolAttr(token.attrs.val)) font.strike = true;
				} else if (name === "color") {
					const color = parseColor(token.attrs);
					if (color !== undefined) font.color = color;
				}
			} else if (name === "fills") {
				if (!token.selfClosing) inFills = true;
			} else if (name === "fill" && inFills) {
				if (token.selfClosing) fills.push(undefined);
				else {
					inFill = true;
					fill = undefined;
				}
			} else if (name === "patternFill" && inFill) {
				const patternType = token.attrs.patternType;
				fill = {
					patternType:
						patternType !== undefined && PATTERN_TYPES.has(patternType as PatternType)
							? (patternType as PatternType)
							: "none",
				};
			} else if (name === "fgColor" && fill !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) fill.fgColor = color;
			} else if (name === "bgColor" && fill !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) fill.bgColor = color;
			} else if (name === "borders") {
				if (!token.selfClosing) inBorders = true;
			} else if (name === "border" && inBorders) {
				if (token.selfClosing) borders.push({});
				else border = {};
			} else if (
				border !== undefined &&
				(name === "left" || name === "right" || name === "top" || name === "bottom")
			) {
				const style = token.attrs.style;
				const lineStyle =
					style !== undefined && BORDER_LINE_STYLES.has(style as BorderLineStyle)
						? (style as BorderLineStyle)
						: undefined;
				if (token.selfClosing) {
					// <left style="thin"/> — complete edge; <left/> — no border on this edge.
					if (lineStyle !== undefined && border !== undefined)
						border[name] = { style: lineStyle };
				} else {
					edgeName = name;
					edgeStyle = lineStyle;
					edgeColor = undefined;
				}
			} else if (name === "color" && edgeName !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) edgeColor = color;
			} else if (name === "cellXfs") {
				if (!token.selfClosing) inCellXfs = true;
			} else if (name === "xf" && inCellXfs) {
				const numFmtId = Number(token.attrs.numFmtId ?? "0");
				const fontId = Number(token.attrs.fontId ?? "0");
				const fillId = Number(token.attrs.fillId ?? "0");
				const borderId = Number(token.attrs.borderId ?? "0");
				xfs.push({
					numFmtId: Number.isInteger(numFmtId) ? numFmtId : 0,
					fontId: Number.isInteger(fontId) ? fontId : 0,
					fillId: Number.isInteger(fillId) ? fillId : 0,
					borderId: Number.isInteger(borderId) ? borderId : 0,
					alignment: undefined,
					protection: undefined,
				});
				if (!token.selfClosing) inXf = true;
			} else if (name === "alignment" && inXf) {
				// Inline alignment belongs to the last-opened cellXfs <xf>.
				const alignment = parseAlignment(token.attrs);
				if (alignment !== undefined && xfs.length > 0) {
					const last = xfs[xfs.length - 1] as XfRecord;
					xfs[xfs.length - 1] = { ...last, alignment };
				}
			} else if (name === "protection" && inXf) {
				// Inline <protection locked hidden> belongs to the last-opened cellXfs <xf> (F10.3). The
				// `inXf` gate scopes it to cellXfs — a <dxf>'s own <protection> never reaches here.
				const protection = parseProtection(token.attrs);
				if (protection !== undefined && xfs.length > 0) {
					const last = xfs[xfs.length - 1] as XfRecord;
					xfs[xfs.length - 1] = { ...last, protection };
				}
			}
		} else if (token.kind === "close") {
			// Section closes FLUSH any dangling builder (a record whose own close tag never came —
			// misnested input the non-validating tokenizer passes through). Flushing rather than
			// discarding keeps the table index aligned with what the producer wrote, and clearing
			// the builder stops it from capturing look-alike elements later in the document (a
			// <dxf> block legitimately contains <font>/<fill>/<border> children — adversarial
			// review showed a leaked builder grafting a dxf fill/color onto cell styles).
			if (name === "numFmts") inNumFmts = false;
			else if (name === "fonts") {
				inFonts = false;
				if (font !== undefined) {
					fonts.push(font);
					font = undefined;
				}
			} else if (name === "font") {
				if (font !== undefined) {
					fonts.push(font);
					font = undefined;
				}
			} else if (name === "fills") {
				inFills = false;
				if (inFill) {
					fills.push(fill);
					fill = undefined;
					inFill = false;
				}
			} else if (name === "fill") {
				if (inFill) {
					fills.push(fill);
					fill = undefined;
					inFill = false;
				}
			} else if (name === "borders") {
				inBorders = false;
				if (border !== undefined) {
					commitEdge();
					borders.push(border);
					border = undefined;
				}
			} else if (name === "border") {
				if (border !== undefined) {
					commitEdge();
					borders.push(border);
					border = undefined;
				}
			} else if (name === "left" || name === "right" || name === "top" || name === "bottom") {
				if (edgeName === name) commitEdge();
			} else if (name === "cellXfs") {
				inCellXfs = false;
				inXf = false;
			} else if (name === "xf") inXf = false;
		}
	}

	function isDateStyle(styleIndex: number | undefined): boolean {
		// An omitted `s` means style 0, the implicit default format.
		const numFmtId = xfs[styleIndex ?? 0]?.numFmtId;
		if (numFmtId === undefined) return false;
		const custom = customFormats.get(numFmtId);
		return custom !== undefined ? isDateFormatCode(custom) : isBuiltinDateId(numFmtId);
	}

	function formatCode(styleIndex: number | undefined): string | undefined {
		// An omitted `s` means style 0, the implicit default format. A custom code for the id
		// wins over the built-in table (a file may redefine one); unknown ids stay undefined.
		const numFmtId = xfs[styleIndex ?? 0]?.numFmtId;
		if (numFmtId === undefined) return undefined;
		return customFormats.get(numFmtId) ?? BUILTIN_FORMATS[numFmtId];
	}

	// Materialize (and cache) the resolved CellStyle for an xf index. The cache makes results
	// reference-stable — cellStyle(i) === cellStyle(i) — which downstream consumers (the bridge,
	// the writer's interner) rely on to dedup styles by identity instead of deep comparison.
	const styleCache = new Map<number, CellStyle | undefined>();

	function cellStyle(styleIndex: number | undefined): CellStyle | undefined {
		const index = styleIndex ?? 0;
		const cached = styleCache.get(index);
		if (cached !== undefined || styleCache.has(index)) return cached;

		const xf = xfs[index];
		let result: CellStyle | undefined;
		if (xf !== undefined) {
			const style: {
				numberFormat?: string;
				font?: FontStyle;
				fill?: FillStyle;
				border?: BorderStyle;
				alignment?: Alignment;
				protection?: CellProtection;
			} = {};
			// numberFormat: "General" is the absence of a format, not a format — whether as id 0
			// or as an explicitly interned custom code (LibreOffice writes <numFmt numFmtId="164"
			// formatCode="General"/> and points xfs at it). Normalizing it away here keeps the
			// style model symmetric with the writer, which reverse-maps 'General' to id 0. Locale
			// ids with no portable code stay absent too (documented; nothing faithful to return).
			if (xf.numFmtId !== 0) {
				const code = customFormats.get(xf.numFmtId) ?? BUILTIN_FORMATS[xf.numFmtId];
				// An EMPTY custom code is as meaningless as General — degrade both to "no format".
				// An XML-UNSAFE code (a control character can arrive via a decoded numeric
				// reference in formatCode) is unwritable, exactly like an unsafe font name —
				// degrade it too, or the bridge would crash on a file the reader tolerated.
				if (code !== undefined && code !== "" && code !== "General" && isXmlSafe(code)) {
					style.numberFormat = code;
				}
			}
			// font: id 0 is the workbook default font — "no per-cell font", by definition. A
			// non-zero id counts even if its content happens to match the default (content dedup
			// is the writer's job, not the reader's).
			if (xf.fontId !== 0) {
				const fontRecord = fonts[xf.fontId];
				if (fontRecord !== undefined && hasKeys(fontRecord)) style.font = fontRecord;
			}
			// fill: judged by VALUE, not id — fill 0 is 'none' by spec, and any other id whose
			// pattern is 'none' paints nothing either.
			const fillRecord = fills[xf.fillId];
			if (fillRecord !== undefined && fillRecord.patternType !== "none")
				style.fill = fillRecord;
			// border: only when at least one edge draws a line.
			const borderRecord = borders[xf.borderId];
			if (borderRecord !== undefined && hasKeys(borderRecord)) style.border = borderRecord;
			if (xf.alignment !== undefined) style.alignment = xf.alignment;
			if (xf.protection !== undefined) style.protection = xf.protection;
			result = hasKeys(style) ? style : undefined;
		}
		styleCache.set(index, result);
		return result;
	}

	return { isDateStyle, formatCode, cellStyle };
}
