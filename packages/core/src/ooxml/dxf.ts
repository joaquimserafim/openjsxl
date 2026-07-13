import type {
	Alignment,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	Color,
	DxfFill,
	DxfStyle,
	FontStyle,
	PatternType,
	UnderlineStyle,
} from "../types";
import { isXmlSafe, localName } from "../utils";
import { tokenize } from "../xml";
import { BORDER_LINE_STYLES, boolAttr, PATTERN_TYPES, parseAlignment, parseColor } from "./styles";

// Differential-style (`<dxfs>`) parser (F9.3). A conditional-formatting rule (or a table region)
// references a dxf by a numeric `dxfId` — an index into styles.xml's `<dxfs>` table. This reads that
// table into positional {@link DxfStyle}s; a caller resolves `dxfId` against the array and never sees
// the index (decision 3).
//
// A `<dxf>` reuses the same `<font>`/`<border>`/`<alignment>`/`<numFmt>` element shapes as the style
// tables (shared helpers), but its FILL is kept RAW — `patternType` is usually absent and the visible
// color is `bgColor`, the exact inverse of a cell fill. Normalizing it would silently swap every
// highlight color, so nothing is dropped. TOLERANT: unparseable pieces degrade to absent; the array
// stays index-aligned with the file (an empty `<dxf/>` is a real, empty slot). Never throws.
//
// `<protection>` inside a dxf is a NAMED drop (rare in conditional formatting; not modelled).

// The <font> child elements a dxf reads — gated by NAME so a dangling <font> can't swallow structure.
const FONT_CHILDREN = new Set(["name", "sz", "b", "i", "u", "strike", "color"]);

interface FontBuilder {
	name?: string;
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: UnderlineStyle;
	strike?: boolean;
	color?: Color;
}

interface DxfBuilder {
	numberFormat?: string;
	font?: FontBuilder;
	fill?: { patternType?: PatternType; fgColor?: Color; bgColor?: Color };
	border?: { top?: BorderEdge; right?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge };
	alignment?: Alignment;
}

const hasKeys = (o: object): boolean => Object.keys(o).length > 0;

/** Parse styles.xml's `<dxfs>` table into positional {@link DxfStyle}s (index = `dxfId`). */
export function parseDxfs(xml: string): DxfStyle[] {
	const dxfs: DxfStyle[] = [];
	let inDxfs = false;
	let dxf: DxfBuilder | undefined; // the <dxf> currently open

	// <font> builder state (only meaningful while a <dxf> is open).
	let inFont = false;
	// <fill> builder state.
	let inFill = false;
	// <border> edge state.
	let edgeName: "top" | "right" | "bottom" | "left" | undefined;
	let edgeStyle: BorderLineStyle | undefined;
	let edgeColor: Color | undefined;
	const commitEdge = (): void => {
		if (dxf?.border !== undefined && edgeName !== undefined && edgeStyle !== undefined) {
			dxf.border[edgeName] =
				edgeColor !== undefined
					? { style: edgeStyle, color: edgeColor }
					: { style: edgeStyle };
		}
		edgeName = undefined;
		edgeStyle = undefined;
		edgeColor = undefined;
	};

	const finishDxf = (): void => {
		if (dxf === undefined) return;
		commitEdge();
		const style: {
			numberFormat?: string;
			font?: FontStyle;
			fill?: DxfFill;
			border?: BorderStyle;
			alignment?: Alignment;
		} = {};
		if (dxf.numberFormat !== undefined) style.numberFormat = dxf.numberFormat;
		if (dxf.font !== undefined && hasKeys(dxf.font)) style.font = dxf.font;
		if (dxf.fill !== undefined && hasKeys(dxf.fill)) style.fill = dxf.fill;
		if (dxf.border !== undefined && hasKeys(dxf.border)) style.border = dxf.border;
		if (dxf.alignment !== undefined) style.alignment = dxf.alignment;
		dxfs.push(style);
		dxf = undefined;
		inFont = false;
		inFill = false;
	};

	for (const token of tokenize(xml)) {
		if (token.kind === "text") continue;
		const name = localName(token.name);

		if (token.kind === "open") {
			if (name === "dxfs") {
				if (!token.selfClosing) inDxfs = true;
			} else if (name === "dxf" && inDxfs) {
				dxf = {};
				inFont = false;
				inFill = false;
				edgeName = undefined;
				if (token.selfClosing) finishDxf();
			} else if (dxf === undefined) {
				// ignore anything outside an open <dxf>
			} else if (name === "font") {
				dxf.font = {};
				inFont = !token.selfClosing;
			} else if (inFont && FONT_CHILDREN.has(name)) {
				const f = dxf.font;
				if (f === undefined) {
					// unreachable: inFont implies dxf.font set
				} else if (name === "name") {
					const val = token.attrs.val;
					if (val !== undefined && val !== "" && isXmlSafe(val)) f.name = val;
				} else if (name === "sz") {
					const size = Number(token.attrs.val);
					if (Number.isFinite(size) && size > 0) f.size = size;
				} else if (name === "b") {
					if (boolAttr(token.attrs.val)) f.bold = true;
				} else if (name === "i") {
					if (boolAttr(token.attrs.val)) f.italic = true;
				} else if (name === "u") {
					const val = token.attrs.val ?? "single";
					if (val === "single" || val === "double") f.underline = val;
				} else if (name === "strike") {
					if (boolAttr(token.attrs.val)) f.strike = true;
				} else if (name === "color") {
					const color = parseColor(token.attrs);
					if (color !== undefined) f.color = color;
				}
			} else if (name === "numFmt") {
				const code = token.attrs.formatCode;
				// Match the cell-style model: carry the CODE string, and only when it is writable.
				if (code !== undefined && code !== "" && isXmlSafe(code)) dxf.numberFormat = code;
			} else if (name === "fill") {
				dxf.fill = {};
				inFill = !token.selfClosing;
			} else if (name === "patternFill" && inFill && dxf.fill !== undefined) {
				// patternType is OPTIONAL in a dxf (raw); keep it only when it is a known value.
				const pt = token.attrs.patternType;
				if (pt !== undefined && PATTERN_TYPES.has(pt as PatternType)) {
					dxf.fill.patternType = pt as PatternType;
				}
			} else if (name === "fgColor" && inFill && dxf.fill !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) dxf.fill.fgColor = color;
			} else if (name === "bgColor" && inFill && dxf.fill !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) dxf.fill.bgColor = color;
			} else if (name === "border") {
				dxf.border = {};
			} else if (
				dxf.border !== undefined &&
				(name === "left" || name === "right" || name === "top" || name === "bottom")
			) {
				const style = token.attrs.style;
				const lineStyle =
					style !== undefined && BORDER_LINE_STYLES.has(style as BorderLineStyle)
						? (style as BorderLineStyle)
						: undefined;
				if (token.selfClosing) {
					if (lineStyle !== undefined) dxf.border[name] = { style: lineStyle };
				} else {
					edgeName = name;
					edgeStyle = lineStyle;
					edgeColor = undefined;
				}
			} else if (name === "color" && edgeName !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) edgeColor = color;
			} else if (name === "alignment") {
				const alignment = parseAlignment(token.attrs);
				if (alignment !== undefined) dxf.alignment = alignment;
			}
		} else if (token.kind === "close") {
			if (name === "dxfs") inDxfs = false;
			else if (name === "dxf") finishDxf();
			else if (name === "font") inFont = false;
			else if (name === "fill") inFill = false;
			else if (name === "left" || name === "right" || name === "top" || name === "bottom") {
				if (edgeName === name) commitEdge();
			}
		}
	}
	// A <dxfs> that never closed (truncated file) still flushes its open dxf.
	finishDxf();
	return dxfs;
}
