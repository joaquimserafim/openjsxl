import type { AnchorPoint } from "../types";
import { localName } from "../utils";
import { tokenize } from "../xml";
import { MAX_COL, MAX_ROW } from "./a1";

// The largest legal EMU value (2³¹−1) for anchor offsets and extents — a SHARED bound: the tolerant
// reader CLAMPS a producer's out-of-range value into 0..MAX_EMU here, and the writer REJECTS an
// out-of-range input typed (writer/sheet.ts emuValue) — both reading this one constant, so whatever
// the reader returns is always writable (the a1.ts MAX_ROW/MAX_COL pattern).
export const MAX_EMU = 0x7fffffff;

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
	// Pull a producer's value back into legal range. The tolerant reader CLAMPS out-of-bounds
	// numbers (a negative offset, a 2³² extent, a cell beyond Excel's grid) instead of returning
	// them verbatim — the writer would refuse them typed, and one malformed picture must not make
	// a whole file un-rewritable. Same shared-bounds contract as column widths and row heights.
	const clamp = (n: number, lo: number, hi: number): number => (n < lo ? lo : n > hi ? hi : n);
	const finish = (p: MutablePoint): AnchorPoint => ({
		col: clamp(p.col, 0, MAX_COL - 1) + 1, // OOXML anchors are 0-based; the public model is 1-based
		row: clamp(p.row, 0, MAX_ROW - 1) + 1,
		colOff: clamp(p.colOff, 0, MAX_EMU),
		rowOff: clamp(p.rowOff, 0, MAX_EMU),
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
				// The extent is clamped into legal EMU range like every other anchor number.
				ext = {
					cx: clamp(toInt(token.attrs.cx), 0, MAX_EMU),
					cy: clamp(toInt(token.attrs.cy), 0, MAX_EMU),
				};
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
				//
				// The anchor ELEMENT names its shape: a twoCellAnchor is positioned by <to>, a
				// oneCellAnchor by a fixed <ext>. Take only the field the kind calls for — so a
				// malformed anchor carrying BOTH (e.g. a stray <ext> before a two-cell <pic>)
				// re-emits as exactly one, and the reader keeps its promise that every image it
				// returns is valid writer input (the writer requires exactly one of to/ext). If the
				// field the kind needs is absent, the picture can't be placed — drop it, the same
				// degradation as a missing blip or media part.
				const anchorTo = kind === "twoCellAnchor" ? to : undefined;
				const anchorExt = kind === "oneCellAnchor" ? ext : undefined;
				if (
					hasPic &&
					!hasGroup &&
					embed &&
					from !== undefined &&
					(anchorTo === undefined) !== (anchorExt === undefined)
				) {
					images.push({
						embed,
						...(picName !== undefined && picName !== "" ? { name: picName } : {}),
						from: finish(from),
						...(anchorTo !== undefined ? { to: finish(anchorTo) } : {}),
						...(anchorExt !== undefined ? { ext: anchorExt } : {}),
						...(editAs !== undefined ? { editAs } : {}),
					});
				}
				reset();
			}
		}
	}
	return images;
}

// ── Media types ──────────────────────────────────────────────────────────────────────────────
// ONE canonical mime ↔ extension source. The writer's allowlist (below) is the master; the
// reader's ext → mime map DERIVES its overlapping entries from it, and the content-types emitter
// (writer/parts.ts) derives its Default entries the same way — three former hand-synced copies
// that could drift are now one definition.

/**
 * The image media types the WRITER accepts, mapped to the media-part extension it emits
 * (`xl/media/imageN.<ext>`). This is the FULL read set (0.6): every type the reader can report
 * from a real file is writable, so any file with pictures round-trips — the writer never decodes
 * the bytes, it only re-embeds what a producer embedded. Only a genuinely unknown type (the
 * reader's `application/octet-stream` fallback) has no part name to emit and refuses typed.
 */
export const MEDIA_MIME_TO_EXT: Readonly<Record<string, string>> = {
	"image/png": "png",
	"image/jpeg": "jpeg",
	"image/gif": "gif",
	"image/bmp": "bmp",
	"image/tiff": "tiff",
	"image/webp": "webp",
	"image/x-emf": "emf",
	"image/x-wmf": "wmf",
};

// Media type from a part's file extension — the read side, DERIVED from the allowlist above
// (spelled identically by construction, so the mime the reader reports for an embedded image
// always re-passes the writer's gate). The two extras are alternate SPELLINGS of derived types
// that real producers use for part names, not additional types.
const MIME_BY_EXT: Readonly<Record<string, string>> = {
	...Object.fromEntries(Object.entries(MEDIA_MIME_TO_EXT).map(([mime, ext]) => [ext, mime])),
	jpg: "image/jpeg",
	tif: "image/tiff",
};

/**
 * Guess an image's type from its filename ending — `image1.png` → `image/png`. An unknown ending
 * falls back to the generic `application/octet-stream`.
 */
export function mimeForMediaPath(path: string): string {
	const dot = path.lastIndexOf(".");
	const ext = dot === -1 ? "" : path.slice(dot + 1).toLowerCase();
	// Own-key lookup only: a hostile part name like `image1.constructor` must fall through to the
	// generic type, not surface an Object.prototype member as the "mime" (the reader-side twin of
	// the writer's Object.hasOwn allowlist gate).
	return Object.hasOwn(MIME_BY_EXT, ext)
		? (MIME_BY_EXT[ext] as string)
		: "application/octet-stream";
}
