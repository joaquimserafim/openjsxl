import type { Color } from "../types";
import { localName } from "../utils";
import { tokenize } from "../xml";

// Theme color resolution (F5.3). SpreadsheetML keeps font/fill/border colors as raw {theme, tint?}
// indexes into the workbook theme's <clrScheme>; the reader never resolves them, because the raw
// form is what round-trips faithfully. This module parses that scheme and applies Excel's tint
// algorithm so `Workbook.resolveColor` can hand a consumer a concrete ARGB on request. The stored
// color model is unchanged — resolution is opt-in.

// The 12 <clrScheme> children in the order every theme writes them (DrawingML document order).
const SCHEME_ORDER = [
	"dk1",
	"lt1",
	"dk2",
	"lt2",
	"accent1",
	"accent2",
	"accent3",
	"accent4",
	"accent5",
	"accent6",
	"hlink",
	"folHlink",
] as const;

// A SpreadsheetML `theme="N"` index is NOT the document order above: Excel swaps each dark/light
// pair, so index 0→lt1 (Background 1), 1→dk1 (Text 1), 2→lt2 (Background 2), 3→dk2 (Text 2); the
// accents and links pass straight through. This maps a theme index to its <clrScheme> slot.
const THEME_TO_SCHEME = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11] as const;

// A bare <a:sysClr val> with no lastClr: the only two system colors a theme uses resolve to these.
const SYS_CLR: Readonly<Record<string, string>> = { windowText: "000000", window: "FFFFFF" };

const HEX6 = /^[0-9A-Fa-f]{6}$/;

/**
 * A parsed theme color scheme: 12 uppercase 6-hex-digit RGB strings indexed by the SpreadsheetML
 * theme index, so `colors[4]` is accent1 and `colors[0]` is the light Background-1 color.
 */
export type ThemeColors = readonly string[];

/**
 * Parse `theme1.xml`'s `<a:clrScheme>` into a {@link ThemeColors} table, or `undefined` when the
 * scheme is absent or any of the 12 slots can't be resolved to an RGB (so resolution degrades to
 * "no answer" rather than a wrong one). Each slot is an `<a:srgbClr val>` or an `<a:sysClr lastClr>`
 * (falling back to the system color's standard value when `lastClr` is missing).
 */
export function parseTheme(xml: string): ThemeColors | undefined {
	const doc: (string | undefined)[] = new Array(12).fill(undefined);
	let inScheme = false;
	let slot = -1; // index into SCHEME_ORDER of the child we're inside, or -1
	for (const token of tokenize(xml)) {
		if (token.kind === "open") {
			const name = localName(token.name);
			if (name === "clrScheme") {
				inScheme = true;
			} else if (inScheme && slot === -1) {
				const i = SCHEME_ORDER.indexOf(name as (typeof SCHEME_ORDER)[number]);
				if (i !== -1) slot = i;
			} else if (inScheme && slot !== -1 && doc[slot] === undefined) {
				// The color element inside the current scheme slot — first one wins.
				if (name === "srgbClr") {
					const val = token.attrs.val;
					if (val !== undefined && HEX6.test(val)) doc[slot] = val.toUpperCase();
				} else if (name === "sysClr") {
					const last = token.attrs.lastClr;
					const resolved =
						last !== undefined && HEX6.test(last)
							? last.toUpperCase()
							: SYS_CLR[token.attrs.val ?? ""];
					if (resolved !== undefined) doc[slot] = resolved;
				}
			}
		} else if (token.kind === "close") {
			const name = localName(token.name);
			// Nothing after </clrScheme> matters — the fmtScheme also holds srgbClr elements we must
			// not read as scheme colors, so stop here.
			if (name === "clrScheme") break;
			if (slot !== -1 && name === SCHEME_ORDER[slot]) slot = -1;
		}
	}
	if (doc.some((c) => c === undefined)) return undefined;
	// Reorder into theme-index order so resolveColor(theme=N) is a direct lookup.
	return THEME_TO_SCHEME.map((s) => doc[s] as string);
}

// ── Excel's tint algorithm (Win32 integer HLS, HLSMAX=240) ─────────────────────────────────────
// Excel resolves a {theme, tint} color by converting the base RGB to integer HLS, scaling the
// luminance toward black (tint<0) or white (tint>0), and converting back. Reproduced here to the
// bit: this matches Excel's own swatch values exactly (the float-HSL variant other libraries use is
// off by ~1 per channel). All the divisors below divide 240 evenly, so the integer arithmetic
// mirrors the classic ColorRGBToHLS / ColorHLSToRGB routines.

const HLSMAX = 240;
const RGBMAX = 255;

const idiv = (a: number, b: number): number => Math.floor(a / b);

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
	const cMax = Math.max(r, g, b);
	const cMin = Math.min(r, g, b);
	const sum = cMax + cMin;
	const l = idiv(sum * HLSMAX + RGBMAX, 2 * RGBMAX);
	if (cMax === cMin) return [160, l, 0]; // achromatic — hue is undefined (arbitrary), saturation 0
	const d = cMax - cMin;
	const s =
		l <= HLSMAX / 2
			? idiv(d * HLSMAX + idiv(sum, 2), sum)
			: idiv(d * HLSMAX + idiv(2 * RGBMAX - sum, 2), 2 * RGBMAX - sum);
	const rd = idiv((cMax - r) * (HLSMAX / 6) + idiv(d, 2), d);
	const gd = idiv((cMax - g) * (HLSMAX / 6) + idiv(d, 2), d);
	const bd = idiv((cMax - b) * (HLSMAX / 6) + idiv(d, 2), d);
	let h: number;
	if (cMax === r) h = bd - gd;
	else if (cMax === g) h = HLSMAX / 3 + rd - bd;
	else h = (2 * HLSMAX) / 3 + gd - rd;
	if (h < 0) h += HLSMAX;
	if (h > HLSMAX) h -= HLSMAX;
	return [h, l, s];
}

function hueToRgb(n1: number, n2: number, hue: number): number {
	if (hue < 0) hue += HLSMAX;
	if (hue > HLSMAX) hue -= HLSMAX;
	if (hue < HLSMAX / 6) return n1 + idiv((n2 - n1) * hue + HLSMAX / 12, HLSMAX / 6);
	if (hue < HLSMAX / 2) return n2;
	if (hue < (2 * HLSMAX) / 3)
		return n1 + idiv((n2 - n1) * ((2 * HLSMAX) / 3 - hue) + HLSMAX / 12, HLSMAX / 6);
	return n1;
}

function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
	if (s === 0) {
		const v = idiv(l * RGBMAX + HLSMAX / 2, HLSMAX);
		return [v, v, v];
	}
	const m2 =
		l <= HLSMAX / 2
			? idiv(l * (HLSMAX + s) + HLSMAX / 2, HLSMAX)
			: l + s - idiv(l * s + HLSMAX / 2, HLSMAX);
	const m1 = 2 * l - m2;
	return [
		idiv(hueToRgb(m1, m2, h + HLSMAX / 3) * RGBMAX + HLSMAX / 2, HLSMAX),
		idiv(hueToRgb(m1, m2, h) * RGBMAX + HLSMAX / 2, HLSMAX),
		idiv(hueToRgb(m1, m2, h - HLSMAX / 3) * RGBMAX + HLSMAX / 2, HLSMAX),
	];
}

const hex2 = (n: number): string => n.toString(16).padStart(2, "0").toUpperCase();

/**
 * Apply an Excel `tint` (−1…1) to a 6-hex-digit RGB, returning a 6-hex-digit RGB. `tint === 0` is
 * the identity (the integer HLS round trip is lossy, so it must short-circuit); otherwise the
 * luminance is scaled toward black/white and rounded half-up, matching JS `Math.round`.
 */
export function resolveTint(rgb6: string, tint: number): string {
	// A tint outside [-1, 1] (or a non-finite one, e.g. from a malformed file the reader kept) is
	// clamped rather than pushed out of gamut into a garbage RGB; 0 is the identity.
	const t = Number.isFinite(tint) ? Math.max(-1, Math.min(1, tint)) : 0;
	if (t === 0) return rgb6.toUpperCase();
	const r = Number.parseInt(rgb6.slice(0, 2), 16);
	const g = Number.parseInt(rgb6.slice(2, 4), 16);
	const b = Number.parseInt(rgb6.slice(4, 6), 16);
	const [h, l, s] = rgbToHls(r, g, b);
	const scaled = t < 0 ? l * (1 + t) : l * (1 - t) + (HLSMAX - HLSMAX * (1 - t));
	const [nr, ng, nb] = hlsToRgb(h, Math.round(scaled), s);
	return `${hex2(nr)}${hex2(ng)}${hex2(nb)}`;
}

// An rgb value from the reader is HEX_COLOR-validated (6 or 8 hex, any case). Normalize to an
// 8-digit uppercase ARGB — a 6-digit value is opaque, so it gains an FF alpha.
function normalizeArgb(rgb: string): string {
	const up = rgb.toUpperCase();
	return up.length === 6 ? `FF${up}` : up;
}

/**
 * Resolve a raw {@link Color} to an 8-digit ARGB string, or `undefined` when it can't be resolved:
 * an `{auto}` color (the consumer decides — usually black), an `{indexed}` palette color (not
 * resolved in F5.3 — the raw index is preserved), or a `{theme}` color when no theme table exists
 * or the index is out of range.
 */
export function resolveColor(color: Color, theme: ThemeColors | undefined): string | undefined {
	if ("rgb" in color) return normalizeArgb(color.rgb);
	if ("theme" in color) {
		if (theme === undefined) return undefined;
		const base = theme[color.theme];
		if (base === undefined) return undefined;
		return `FF${resolveTint(base, color.tint ?? 0)}`;
	}
	return undefined;
}
