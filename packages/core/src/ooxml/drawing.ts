import type { AnchorPoint } from "../types";
import { localName } from "../utils";
import { tokenize } from "../xml";

// Parse a spreadsheetDrawing part (xl/drawings/drawingN.xml) into the picture anchors it declares,
// walking the SAX tokenizer rather than building a DOM (F6.2). A drawing holds a sequence of
// <xdr:oneCellAnchor> / <xdr:twoCellAnchor> / <xdr:absoluteAnchor>, each wrapping ONE drawing
// object. We keep only PICTURE anchors (those containing an <xdr:pic> with an <a:blip r:embed>);
// shapes, charts, and group/graphic frames carry no embedded image and are skipped, as are
// absolute-anchored pictures (their geometry isn't cell-relative — out of the shared model).
//
// This module is pure: it returns the unresolved r:embed id for each picture; the reader resolves
// that through the drawing's own relationships to a media part and reads the bytes (the I/O lives
// in the reader, matching every other ooxml parser).

/** One picture anchor from a drawing part — geometry + the unresolved blip relationship id. */
export interface DrawingImage {
	/** The `r:embed` relationship id on the picture's `<a:blip>` (resolved by the reader). */
	readonly embed: string;
	/** The picture's `<xdr:cNvPr name>` when present. */
	readonly name?: string;
	/** Top-left anchor cell (1-based) + EMU offset. */
	readonly from: AnchorPoint;
	/** Bottom-right anchor cell (1-based) — present for a two-cell anchor. */
	readonly to?: AnchorPoint;
	/** Fixed EMU extent — present for a one-cell anchor. */
	readonly ext?: { readonly cx: number; readonly cy: number };
	/** The anchor's move/size behaviour, when the producer declared `editAs`. */
	readonly editAs?: "twoCell" | "oneCell" | "absolute";
}

const ANCHOR_KINDS = new Set(["oneCellAnchor", "twoCellAnchor", "absoluteAnchor"]);

// A cell-position field being assembled from its text content.
type PointField = "col" | "colOff" | "row" | "rowOff";
interface MutablePoint {
	col: number;
	row: number;
	colOff: number;
	rowOff: number;
}

// Read a whole number from XML text. Anything missing or not a number becomes 0, so a broken file
// can't crash us.
function toInt(s: string | undefined): number {
	if (s === undefined) return 0;
	const n = Number(s);
	return Number.isFinite(n) ? Math.trunc(n) : 0;
}

// Find which relationship id points at this picture's image file. It's usually written `r:embed`,
// but the `r:` prefix isn't guaranteed, so also accept any attribute ending in `embed`.
function blipEmbed(attrs: Readonly<Record<string, string>>): string | undefined {
	if (attrs["r:embed"] !== undefined) return attrs["r:embed"];
	for (const key of Object.keys(attrs)) {
		if (localName(key) === "embed") return attrs[key];
	}
	return undefined;
}

/** Parse a drawing part into its picture anchors, in document order. Never throws. */
export function parseDrawing(xml: string): DrawingImage[] {
	const images: DrawingImage[] = [];

	// State for the anchor currently being assembled.
	let kind: string | undefined; // set while inside an anchor element
	let editAs: "twoCell" | "oneCell" | "absolute" | undefined;
	let from: MutablePoint | undefined;
	let to: MutablePoint | undefined;
	let ext: { cx: number; cy: number } | undefined;
	let embed: string | undefined;
	let picName: string | undefined;
	let hasPic = false;
	// A picture nested in a <grpSp> group carries the GROUP's anchor, not its own, so a grouped
	// picture is skipped — matching this module's "group frames skipped" contract.
	let hasGroup = false;
	// Which <xdr:from>/<xdr:to> point is open, and which of its fields is capturing text.
	let point: MutablePoint | undefined;
	let field: PointField | undefined;
	let fieldText = "";

	const reset = (): void => {
		kind = undefined;
		editAs = undefined;
		from = undefined;
		to = undefined;
		ext = undefined;
		embed = undefined;
		picName = undefined;
		hasPic = false;
		hasGroup = false;
		point = undefined;
		field = undefined;
		fieldText = "";
	};

	const newPoint = (): MutablePoint => ({ col: 0, row: 0, colOff: 0, rowOff: 0 });
	const finish = (p: MutablePoint): AnchorPoint => ({
		col: p.col + 1, // OOXML anchors are 0-based; the public model is 1-based
		row: p.row + 1,
		colOff: p.colOff,
		rowOff: p.rowOff,
	});

	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			if (ANCHOR_KINDS.has(name)) {
				reset();
				kind = name;
				const ea = token.attrs.editAs;
				if (ea === "twoCell" || ea === "oneCell" || ea === "absolute") editAs = ea;
			} else if (kind === undefined) {
				// Ignore anything outside an anchor (wsDr root, extLst, etc.).
			} else if (name === "from") {
				from = newPoint();
				if (!token.selfClosing) point = from;
			} else if (name === "to") {
				to = newPoint();
				if (!token.selfClosing) point = to;
			} else if (
				point !== undefined &&
				(name === "col" || name === "colOff" || name === "row" || name === "rowOff")
			) {
				field = name;
				fieldText = "";
			} else if (
				name === "ext" &&
				!hasPic &&
				ext === undefined &&
				token.attrs.cx !== undefined
			) {
				// The anchor's OWN <ext> (one-cell display size) is a direct child, before <pic> — so
				// only capture an <ext> seen BEFORE the picture. This skips the picture's
				// spPr/a:xfrm/a:ext (the shape-transform extent Excel/LibreOffice emit, which would
				// otherwise overwrite the anchor size) and an extLst <ext> (which carries a uri, no cx).
				ext = { cx: toInt(token.attrs.cx), cy: toInt(token.attrs.cy) };
			} else if (name === "grpSp") {
				hasGroup = true;
			} else if (name === "pic") {
				hasPic = true;
			} else if (name === "cNvPr" && hasPic && picName === undefined) {
				// The picture's own name lives in <pic><nvPicPr><cNvPr> — take the first cNvPr AFTER
				// <pic> opens, so a wrapping group/frame's cNvPr can't masquerade as the picture name.
				picName = token.attrs.name;
			} else if (name === "blip" && embed === undefined) {
				embed = blipEmbed(token.attrs);
			}
		} else if (token.kind === "text") {
			if (field !== undefined) fieldText += token.value;
		} else {
			const name = localName(token.name);
			if (field !== undefined && name === field) {
				if (point !== undefined) point[field] = toInt(fieldText);
				field = undefined;
				fieldText = "";
			} else if (name === "from" || name === "to") {
				point = undefined;
			} else if (ANCHOR_KINDS.has(name) && kind !== undefined) {
				// Keep only cell-anchored pictures with a resolvable blip; skip absoluteAnchor, any
				// anchor whose object isn't a picture (no <xdr:pic>/<a:blip>), and grouped pictures
				// (which carry the group's anchor, not their own).
				if (
					kind !== "absoluteAnchor" &&
					hasPic &&
					!hasGroup &&
					embed &&
					from !== undefined
				) {
					images.push({
						embed,
						...(picName !== undefined && picName !== "" ? { name: picName } : {}),
						from: finish(from),
						...(to !== undefined ? { to: finish(to) } : {}),
						...(ext !== undefined ? { ext } : {}),
						...(editAs !== undefined ? { editAs } : {}),
					});
				}
				reset();
			}
		}
	}
	return images;
}

// Media type from a part's file extension. Content-types Default entries are themselves keyed by
// extension, and our writer (F6.3) sets them from the same map, so this round-trips exactly; an
// unknown extension degrades to a generic binary type rather than guessing.
const MIME_BY_EXT: Readonly<Record<string, string>> = {
	png: "image/png",
	jpg: "image/jpeg",
	jpeg: "image/jpeg",
	gif: "image/gif",
	bmp: "image/bmp",
	tif: "image/tiff",
	tiff: "image/tiff",
	webp: "image/webp",
	emf: "image/x-emf",
	wmf: "image/x-wmf",
};

/**
 * Guess an image's type from its filename ending — `image1.png` → `image/png`. An unknown ending
 * falls back to the generic `application/octet-stream`.
 */
export function mimeForMediaPath(path: string): string {
	const dot = path.lastIndexOf(".");
	const ext = dot === -1 ? "" : path.slice(dot + 1).toLowerCase();
	return MIME_BY_EXT[ext] ?? "application/octet-stream";
}
