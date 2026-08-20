import {
	MAX_PAGE_MARGIN,
	MAX_PAGE_SCALE,
	MAX_PAGE_SETUP_UINT,
	MAX_SPIN_COUNT,
	MIN_PAGE_SCALE,
} from "../ooxml";
import type { HeaderFooter, PageMargins, PageSetup, PrintOptions, SheetProtection } from "../types";

// Pure conversions from worksheet XML attributes to reader model values. Keeping these separate
// lets protection, print setup, and geometry share the same coercion rules without XML parser state.

export type Mutable<T> = { -readonly [K in keyof T]: T[K] };

/** Input `"1"`/`"true"`; output `true`. Input `"0"`/`"false"`; output `false`. */
export function booleanAttribute(value: string | undefined): boolean | undefined {
	if (value === "1" || value === "true") return true;
	if (value === "0" || value === "false") return false;
	return undefined;
}

/** Input `"true"`; output `true`. Input `"0"`, malformed, or absent; output `false`. */
export function trueAttribute(value: string | undefined): boolean {
	return value === "1" || value === "true";
}

/** Input `"12"`, max `10`; output `10`. Input `"NaN"`; output `undefined`. */
export function clampedDoubleAttribute(value: string | undefined, max: number): number | undefined {
	if (value === undefined) return undefined;
	const number = Number(value);
	if (!Number.isFinite(number)) return undefined;
	return number < 0 ? 0 : number > max ? max : number;
}

/** Input `"42"`; output `42`. Input `"1e3"` or a value above `max`; output `undefined`. */
export function unsignedIntAttribute(
	value: string | undefined,
	max = MAX_PAGE_SETUP_UINT,
): number | undefined {
	if (value === undefined || !/^[0-9]+$/.test(value)) return undefined;
	const number = Number(value);
	return Number.isInteger(number) && number <= max ? number : undefined;
}

/** Input `"portrait"` with `["portrait", "landscape"]`; output `"portrait"`. */
export function enumAttribute<T extends string>(
	value: string | undefined,
	allowed: readonly T[],
): T | undefined {
	return value !== undefined && (allowed as readonly string[]).includes(value)
		? (value as T)
		: undefined;
}

const SHEET_PROTECTION_FLAGS = [
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

/** Input `{ sheet: "1", formatCells: "0" }`; output `{ sheet: true, formatCells: false }`. */
export function sheetProtectionFromAttributes(
	attrs: Readonly<Record<string, string>>,
): SheetProtection | undefined {
	const out: Mutable<SheetProtection> = {};
	for (const flag of SHEET_PROTECTION_FLAGS) {
		const value = booleanAttribute(attrs[flag]);
		if (value !== undefined) out[flag] = value;
	}
	if (attrs.password !== undefined) out.password = attrs.password;
	if (attrs.algorithmName !== undefined) out.algorithmName = attrs.algorithmName;
	if (attrs.hashValue !== undefined) out.hashValue = attrs.hashValue;
	if (attrs.saltValue !== undefined) out.saltValue = attrs.saltValue;
	const spinCount = unsignedIntAttribute(attrs.spinCount, MAX_SPIN_COUNT);
	if (spinCount !== undefined) out.spinCount = spinCount;
	return Object.keys(out).length > 0 ? out : undefined;
}

/** Input `{ gridLines: "1" }`; output `{ gridLines: true }`. */
export function printOptionsFromAttributes(
	attrs: Readonly<Record<string, string>>,
): PrintOptions | undefined {
	const out: Mutable<PrintOptions> = {};
	for (const flag of [
		"gridLines",
		"headings",
		"horizontalCentered",
		"verticalCentered",
	] as const) {
		const value = booleanAttribute(attrs[flag]);
		if (value !== undefined) out[flag] = value;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

/** Input all six margin strings; output six numbers. Missing one margin returns `undefined`. */
export function pageMarginsFromAttributes(
	attrs: Readonly<Record<string, string>>,
): PageMargins | undefined {
	const left = clampedDoubleAttribute(attrs.left, MAX_PAGE_MARGIN);
	const right = clampedDoubleAttribute(attrs.right, MAX_PAGE_MARGIN);
	const top = clampedDoubleAttribute(attrs.top, MAX_PAGE_MARGIN);
	const bottom = clampedDoubleAttribute(attrs.bottom, MAX_PAGE_MARGIN);
	const header = clampedDoubleAttribute(attrs.header, MAX_PAGE_MARGIN);
	const footer = clampedDoubleAttribute(attrs.footer, MAX_PAGE_MARGIN);
	if (
		left === undefined ||
		right === undefined ||
		top === undefined ||
		bottom === undefined ||
		header === undefined ||
		footer === undefined
	) {
		return undefined;
	}
	return { left, right, top, bottom, header, footer };
}

/** Input `{ orientation: "landscape", fitToWidth: "1" }`; output the typed pair. */
export function pageSetupFromAttributes(
	attrs: Readonly<Record<string, string>>,
): PageSetup | undefined {
	const out: Mutable<PageSetup> = {};
	const paperSize = unsignedIntAttribute(attrs.paperSize);
	if (paperSize !== undefined) out.paperSize = paperSize;
	const orientation = enumAttribute(attrs.orientation, [
		"default",
		"portrait",
		"landscape",
	] as const);
	if (orientation !== undefined) out.orientation = orientation;
	if (attrs.scale !== undefined && /^[0-9]+$/.test(attrs.scale)) {
		const scale = Number(attrs.scale);
		if (Number.isInteger(scale)) {
			out.scale =
				scale < MIN_PAGE_SCALE
					? MIN_PAGE_SCALE
					: scale > MAX_PAGE_SCALE
						? MAX_PAGE_SCALE
						: scale;
		}
	}
	const fitToWidth = unsignedIntAttribute(attrs.fitToWidth);
	if (fitToWidth !== undefined) out.fitToWidth = fitToWidth;
	const fitToHeight = unsignedIntAttribute(attrs.fitToHeight);
	if (fitToHeight !== undefined) out.fitToHeight = fitToHeight;
	const firstPageNumber = unsignedIntAttribute(attrs.firstPageNumber);
	if (firstPageNumber !== undefined) out.firstPageNumber = firstPageNumber;
	const useFirstPageNumber = booleanAttribute(attrs.useFirstPageNumber);
	if (useFirstPageNumber !== undefined) out.useFirstPageNumber = useFirstPageNumber;
	const blackAndWhite = booleanAttribute(attrs.blackAndWhite);
	if (blackAndWhite !== undefined) out.blackAndWhite = blackAndWhite;
	const draft = booleanAttribute(attrs.draft);
	if (draft !== undefined) out.draft = draft;
	const cellComments = enumAttribute(attrs.cellComments, [
		"none",
		"asDisplayed",
		"atEnd",
	] as const);
	if (cellComments !== undefined) out.cellComments = cellComments;
	const pageOrder = enumAttribute(attrs.pageOrder, ["downThenOver", "overThenDown"] as const);
	if (pageOrder !== undefined) out.pageOrder = pageOrder;
	return Object.keys(out).length > 0 ? out : undefined;
}

export const HEADER_FOOTER_CHILDREN = [
	"oddHeader",
	"oddFooter",
	"evenHeader",
	"evenFooter",
	"firstHeader",
	"firstFooter",
] as const;

export type HeaderFooterChild = (typeof HEADER_FOOTER_CHILDREN)[number];

/** Input `"oddHeader"`; output `true`. Input `"unknown"`; output `false`. */
export function isHeaderFooterChild(name: string): name is HeaderFooterChild {
	return (HEADER_FOOTER_CHILDREN as readonly string[]).includes(name);
}

/** Input `{ differentFirst: "1" }`; output `{ differentFirst: true }`. */
export function headerFooterFromAttributes(
	attrs: Readonly<Record<string, string>>,
): Mutable<HeaderFooter> {
	const out: Mutable<HeaderFooter> = {};
	for (const flag of [
		"differentOddEven",
		"differentFirst",
		"scaleWithDoc",
		"alignWithMargins",
	] as const) {
		const value = booleanAttribute(attrs[flag]);
		if (value !== undefined) out[flag] = value;
	}
	return out;
}
