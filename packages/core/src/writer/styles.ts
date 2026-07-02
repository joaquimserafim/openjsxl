import { XlsxError } from "../errors"
import {
	BORDER_LINE_STYLES,
	BUILTIN_FORMATS,
	H_ALIGNMENTS,
	HEX_COLOR,
	MAX_COLOR_INDEX,
	MAX_INDENT,
	PATTERN_TYPES,
	V_ALIGNMENTS,
} from "../ooxml/styles"
import type {
	Alignment,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	CellStyle,
	Color,
	FillStyle,
	FontStyle,
	HorizontalAlignment,
	PatternType,
	VerticalAlignment,
} from "../types"
import { escapeAttr, isXmlSafe } from "./xml"

// The writer's style registry (F4.2): validates CellStyle input, interns each distinct component
// (font / fill / border / alignment) and cell format (xf) STRUCTURALLY — identical styles collapse
// to one slot no matter how many cells carry them, or whether the caller shared or inlined the
// objects — and emits xl/styles.xml.
//
// Byte-compatibility is a hard contract here: for input with no styles, the registry's tables
// stay at their preseeded defaults and stylesXml() reproduces the pre-F4.2 hardcoded stylesheet
// BYTE-FOR-BYTE (golden-pinned in tests). The preseeds are also Excel's structural invariants —
// files missing them silently misrender or repair:
//
//   font 0    the workbook default font (Calibri 11) — "no per-cell font"
//   fill 0    patternType "none"   ─ reserved pair Excel expects at
//   fill 1    patternType "gray125"─ indexes 0 and 1, in that order
//   border 0  the empty border
//   xf 0      the default cell format (General, all component ids 0)
//   plus one cellStyleXfs entry and the Normal cellStyle, emitted statically.
//
// Validation is strict and names the offending cell: unknown keys, bad enum values, and malformed
// colors throw XlsxError('invalid-input') rather than write a file Excel repairs. The one
// deliberate exception to strictness: `false` booleans, indent 0, and textRotation 0 are DEFAULTS,
// not errors — they normalize away so `{ bold: false }` interns identically to `{}`.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// The built-in date numFmt a bare Date cell gets (mm-dd-yy); the reader maps it back to `date`.
const DATE_NUMFMT_ID = 14

// Reverse of the reader's BUILTIN_FORMATS: an EXACT code match reuses the built-in id and needs
// no <numFmts> entry ('0.00%' → 10). Anything else interns as a custom format from 164 up —
// below 164 is reserved for built-ins. First-wins on the (currently absent) chance of duplicate
// codes in the table, so the mapping is deterministic.
const BUILTIN_CODE_TO_ID: ReadonlyMap<string, number> = (() => {
	const map = new Map<string, number>()
	for (const [id, code] of Object.entries(BUILTIN_FORMATS)) {
		if (!map.has(code)) map.set(code, Number(id))
	}
	return map
})()
const CUSTOM_NUMFMT_BASE = 164

function invalid(ref: string, message: string): never {
	throw new XlsxError("invalid-input", `cell ${ref}: ${message}`)
}

// STRICTLY plain objects: prototype null or Object.prototype. A Map/Set/Date/class instance has
// no own enumerable keys, so it would sail through checkKeys while property access then walks its
// PROTOTYPE (adversarial review: a Map's .size getter validated as font size 2) — reject the
// whole shape instead.
function isPlainObject(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false
	const proto = Object.getPrototypeOf(value)
	return proto === null || proto === Object.prototype
}

function checkKeys(
	ref: string,
	what: string,
	obj: Record<string, unknown>,
	allowed: readonly string[],
): void {
	for (const key of Object.keys(obj)) {
		if (!allowed.includes(key)) invalid(ref, `${what} has an unknown property "${key}"`)
	}
}

// ── Validation + normalization ─────────────────────────────────────────────────────────────────
// Each validator returns a NORMALIZED copy (defaults dropped, key order fixed) or undefined when
// the component normalizes to nothing. The normalized object is both the intern identity (its
// canonical key) and the emission source.

// NOTE for every validator below: each input property is read EXACTLY ONCE into a local, and the
// normalized object is built from the locals. A getter (or Proxy) could otherwise return a valid
// value for the check reads and something else — up to raw markup — for the emission read
// (adversarial review demonstrated attribute injection into styles.xml that way).
function validateColor(ref: string, what: string, raw: unknown): Color {
	if (!isPlainObject(raw)) invalid(ref, `${what} must be a color object`)
	if ("rgb" in raw) {
		checkKeys(ref, what, raw, ["rgb"])
		const rgb = raw.rgb
		if (typeof rgb !== "string" || !HEX_COLOR.test(rgb)) {
			invalid(ref, `${what}.rgb must be 6- or 8-digit hex (got ${JSON.stringify(rgb)})`)
		}
		return { rgb }
	}
	if ("theme" in raw) {
		checkKeys(ref, what, raw, ["theme", "tint"])
		const theme = raw.theme
		if (
			typeof theme !== "number" ||
			!Number.isInteger(theme) ||
			theme < 0 ||
			theme > MAX_COLOR_INDEX
		) {
			invalid(ref, `${what}.theme must be an integer between 0 and ${MAX_COLOR_INDEX}`)
		}
		const tint = raw.tint
		if (tint === undefined) return { theme }
		if (typeof tint !== "number" || !Number.isFinite(tint)) {
			invalid(ref, `${what}.tint must be a finite number`)
		}
		return { theme, tint }
	}
	if ("indexed" in raw) {
		checkKeys(ref, what, raw, ["indexed"])
		const indexed = raw.indexed
		if (
			typeof indexed !== "number" ||
			!Number.isInteger(indexed) ||
			indexed < 0 ||
			indexed > MAX_COLOR_INDEX
		) {
			invalid(ref, `${what}.indexed must be an integer between 0 and ${MAX_COLOR_INDEX}`)
		}
		return { indexed }
	}
	if ("auto" in raw) {
		checkKeys(ref, what, raw, ["auto"])
		if (raw.auto !== true) invalid(ref, `${what}.auto must be true`)
		return { auto: true }
	}
	invalid(ref, `${what} needs one of rgb / theme / indexed / auto`)
}

function validateFont(ref: string, raw: unknown): FontStyle | undefined {
	if (!isPlainObject(raw)) invalid(ref, "style.font must be an object")
	checkKeys(ref, "style.font", raw, [
		"name",
		"size",
		"bold",
		"italic",
		"underline",
		"strike",
		"color",
	])
	const out: {
		name?: string
		size?: number
		bold?: boolean
		italic?: boolean
		underline?: "single" | "double"
		strike?: boolean
		color?: Color
	} = {}
	const name = raw.name
	if (name !== undefined) {
		// The one free-form string in the whole style path — gate it like cell strings and sheet
		// names, or a control character / lone surrogate would corrupt styles.xml itself.
		if (typeof name !== "string" || name.length === 0) {
			invalid(ref, "style.font.name must be a non-empty string")
		}
		if (!isXmlSafe(name)) {
			invalid(
				ref,
				"style.font.name contains a character not allowed in XML (a control character or lone surrogate)",
			)
		}
		out.name = name
	}
	const size = raw.size
	if (size !== undefined) {
		if (typeof size !== "number" || !Number.isFinite(size) || size <= 0) {
			invalid(ref, "style.font.size must be a positive number")
		}
		out.size = size
	}
	for (const flag of ["bold", "italic", "strike"] as const) {
		const value = raw[flag]
		if (value !== undefined) {
			if (typeof value !== "boolean") invalid(ref, `style.font.${flag} must be a boolean`)
			if (value) out[flag] = true // false is the default — normalizes away
		}
	}
	const underline = raw.underline
	if (underline !== undefined) {
		if (underline !== "single" && underline !== "double") {
			invalid(
				ref,
				`style.font.underline must be "single" or "double" (accounting variants are not supported)`,
			)
		}
		out.underline = underline
	}
	const color = raw.color
	if (color !== undefined) out.color = validateColor(ref, "style.font.color", color)
	return Object.keys(out).length > 0 ? out : undefined
}

function validateFill(ref: string, raw: unknown): FillStyle | undefined {
	if (!isPlainObject(raw)) invalid(ref, "style.fill must be an object")
	checkKeys(ref, "style.fill", raw, ["patternType", "fgColor", "bgColor"])
	const patternType = raw.patternType
	if (typeof patternType !== "string" || !PATTERN_TYPES.has(patternType as never)) {
		invalid(ref, `style.fill.patternType must be one of the OOXML pattern types`)
	}
	const rawFg = raw.fgColor
	const fgColor =
		rawFg === undefined ? undefined : validateColor(ref, "style.fill.fgColor", rawFg)
	const rawBg = raw.bgColor
	const bgColor =
		rawBg === undefined ? undefined : validateColor(ref, "style.fill.bgColor", rawBg)
	if (patternType === "none") {
		// "none" paints nothing; colors on it would be silently dead — reject rather than pretend.
		if (fgColor !== undefined || bgColor !== undefined) {
			invalid(ref, 'style.fill with patternType "none" cannot carry colors')
		}
		return undefined // no fill at all — interns to the reserved fill 0
	}
	const out: { patternType: PatternType; fgColor?: Color; bgColor?: Color } = {
		patternType: patternType as PatternType,
	}
	if (fgColor !== undefined) out.fgColor = fgColor
	if (bgColor !== undefined) out.bgColor = bgColor
	return out
}

function validateEdge(ref: string, what: string, raw: unknown): BorderEdge {
	if (!isPlainObject(raw)) invalid(ref, `${what} must be an object`)
	checkKeys(ref, what, raw, ["style", "color"])
	const style = raw.style
	if (typeof style !== "string" || !BORDER_LINE_STYLES.has(style as never)) {
		invalid(ref, `${what}.style must be one of the OOXML border line styles`)
	}
	const color = raw.color
	if (color === undefined) return { style: style as BorderLineStyle }
	return {
		style: style as BorderLineStyle,
		color: validateColor(ref, `${what}.color`, color),
	}
}

function validateBorder(ref: string, raw: unknown): BorderStyle | undefined {
	if (!isPlainObject(raw)) invalid(ref, "style.border must be an object")
	// Diagonal borders are deferred (they need diagonalUp/Down flags the model doesn't carry).
	checkKeys(ref, "style.border", raw, ["top", "right", "bottom", "left"])
	const out: { top?: BorderEdge; right?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge } = {}
	for (const edge of ["top", "right", "bottom", "left"] as const) {
		const value = raw[edge]
		if (value !== undefined) out[edge] = validateEdge(ref, `style.border.${edge}`, value)
	}
	return Object.keys(out).length > 0 ? out : undefined
}

function validateAlignment(ref: string, raw: unknown): Alignment | undefined {
	if (!isPlainObject(raw)) invalid(ref, "style.alignment must be an object")
	checkKeys(ref, "style.alignment", raw, [
		"horizontal",
		"vertical",
		"wrapText",
		"shrinkToFit",
		"indent",
		"textRotation",
	])
	const out: {
		horizontal?: HorizontalAlignment
		vertical?: VerticalAlignment
		wrapText?: boolean
		shrinkToFit?: boolean
		indent?: number
		textRotation?: number
	} = {}
	const horizontal = raw.horizontal
	if (horizontal !== undefined) {
		if (typeof horizontal !== "string" || !H_ALIGNMENTS.has(horizontal as never)) {
			invalid(ref, "style.alignment.horizontal is not a valid value")
		}
		out.horizontal = horizontal as HorizontalAlignment
	}
	const vertical = raw.vertical
	if (vertical !== undefined) {
		if (typeof vertical !== "string" || !V_ALIGNMENTS.has(vertical as never)) {
			invalid(ref, "style.alignment.vertical is not a valid value")
		}
		out.vertical = vertical as VerticalAlignment
	}
	for (const flag of ["wrapText", "shrinkToFit"] as const) {
		const value = raw[flag]
		if (value !== undefined) {
			if (typeof value !== "boolean")
				invalid(ref, `style.alignment.${flag} must be a boolean`)
			if (value) out[flag] = true
		}
	}
	const indent = raw.indent
	if (indent !== undefined) {
		if (
			typeof indent !== "number" ||
			!Number.isInteger(indent) ||
			indent < 0 ||
			indent > MAX_INDENT
		) {
			invalid(ref, `style.alignment.indent must be an integer between 0 and ${MAX_INDENT}`)
		}
		if (indent > 0) out.indent = indent
	}
	const textRotation = raw.textRotation
	if (textRotation !== undefined) {
		if (
			typeof textRotation !== "number" ||
			!Number.isInteger(textRotation) ||
			textRotation < 0 ||
			textRotation > 180
		) {
			invalid(ref, "style.alignment.textRotation must be an integer between 0 and 180")
		}
		if (textRotation > 0) out.textRotation = textRotation
	}
	return Object.keys(out).length > 0 ? out : undefined
}

// ── Canonical keys ─────────────────────────────────────────────────────────────────────────────
// The validators build normalized objects with a FIXED key insertion order, so JSON.stringify is
// a stable structural identity: two styles that mean the same thing produce the same key.

const keyOf = (o: object | undefined): string => (o === undefined ? "" : JSON.stringify(o))

// ── XML emission ───────────────────────────────────────────────────────────────────────────────

function colorXml(tag: string, color: Color): string {
	if ("rgb" in color) return `<${tag} rgb="${color.rgb}"/>`
	if ("theme" in color) {
		const tint = color.tint !== undefined ? ` tint="${String(color.tint)}"` : ""
		return `<${tag} theme="${color.theme}"${tint}/>`
	}
	if ("indexed" in color) return `<${tag} indexed="${color.indexed}"/>`
	return `<${tag} auto="1"/>`
}

function fontXml(font: FontStyle): string {
	// Fixed child order (bold/italic/strike/underline, then size, color, name). Font 0 is
	// preseeded as a literal and never passes through here.
	let out = "<font>"
	if (font.bold) out += "<b/>"
	if (font.italic) out += "<i/>"
	if (font.strike) out += "<strike/>"
	if (font.underline !== undefined) {
		out += font.underline === "single" ? "<u/>" : '<u val="double"/>'
	}
	if (font.size !== undefined) out += `<sz val="${String(font.size)}"/>`
	if (font.color !== undefined) out += colorXml("color", font.color)
	if (font.name !== undefined) out += `<name val="${escapeAttr(font.name)}"/>`
	return `${out}</font>`
}

function fillXml(fill: FillStyle): string {
	const fg = fill.fgColor !== undefined ? colorXml("fgColor", fill.fgColor) : ""
	const bg = fill.bgColor !== undefined ? colorXml("bgColor", fill.bgColor) : ""
	if (fg === "" && bg === "") {
		return `<fill><patternFill patternType="${fill.patternType}"/></fill>`
	}
	return `<fill><patternFill patternType="${fill.patternType}">${fg}${bg}</patternFill></fill>`
}

function borderXml(border: BorderStyle): string {
	// Schema order: left, right, top, bottom. Only edges that draw are emitted; an edge with a
	// color wraps it as a child.
	const edge = (tag: string, e: BorderEdge | undefined): string => {
		if (e === undefined) return ""
		if (e.color === undefined) return `<${tag} style="${e.style}"/>`
		return `<${tag} style="${e.style}">${colorXml("color", e.color)}</${tag}>`
	}
	const inner =
		edge("left", border.left) +
		edge("right", border.right) +
		edge("top", border.top) +
		edge("bottom", border.bottom)
	return inner === "" ? "<border/>" : `<border>${inner}</border>`
}

function alignmentXml(alignment: Alignment): string {
	let attrs = ""
	if (alignment.horizontal !== undefined) attrs += ` horizontal="${alignment.horizontal}"`
	if (alignment.vertical !== undefined) attrs += ` vertical="${alignment.vertical}"`
	if (alignment.textRotation !== undefined) attrs += ` textRotation="${alignment.textRotation}"`
	if (alignment.wrapText) attrs += ' wrapText="1"'
	if (alignment.shrinkToFit) attrs += ' shrinkToFit="1"'
	if (alignment.indent !== undefined) attrs += ` indent="${alignment.indent}"`
	return `<alignment${attrs}/>`
}

function colorUsesTheme(color: Color | undefined): boolean {
	return color !== undefined && "theme" in color
}

// ── The registry ───────────────────────────────────────────────────────────────────────────────

interface XfRecord {
	readonly numFmtId: number
	readonly fontId: number
	readonly fillId: number
	readonly borderId: number
	readonly alignment: Alignment | undefined
}

export interface StyleRegistry {
	/**
	 * The cellXfs index for a cell with this (possibly absent) style; `isDate` forces the built-in
	 * date number format so a bare `Date` keeps its pre-F4.2 behavior. Validates the style and
	 * throws `invalid-input` naming `ref` on anything unrepresentable. Index 0 means "default" —
	 * the caller omits the `s` attribute entirely.
	 */
	xfIndexFor(style: CellStyle | undefined, isDate: boolean, ref: string): number
	/** True when any cell interned a non-default format — i.e. styles.xml must be emitted. */
	needed(): boolean
	/** True when any written color is theme-based — i.e. theme1.xml must be emitted. */
	usesTheme(): boolean
	stylesXml(): string
}

export function createStyleRegistry(): StyleRegistry {
	// Component tables, preseeded with the Excel-required entries (see the header comment). Each
	// table pairs the emitted XML per slot with an intern map from canonical key → slot.
	const fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/></font>']
	const fontIndex = new Map<string, number>([
		[keyOf({ name: "Calibri", size: 11 }), 0],
		["", 0],
	])

	const fills: string[] = [
		'<fill><patternFill patternType="none"/></fill>',
		'<fill><patternFill patternType="gray125"/></fill>',
	]
	const fillIndex = new Map<string, number>([
		["", 0],
		[keyOf({ patternType: "gray125" }), 1],
	])

	const borders: string[] = ["<border/>"]
	const borderIndex = new Map<string, number>([["", 0]])

	const xfs: XfRecord[] = [
		{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: undefined },
	]
	const xfIndex = new Map<string, number>([["0/0/0/0/", 0]])

	// Custom number formats (F4.3): code → id from 164 up, deduped, in first-encounter order.
	// Codes exactly matching a built-in reuse its id instead and never appear here.
	const customFormats = new Map<string, number>()

	let themeUsed = false

	const numFmtIdFor = (code: string): number => {
		const builtin = BUILTIN_CODE_TO_ID.get(code)
		if (builtin !== undefined) return builtin
		let id = customFormats.get(code)
		if (id === undefined) {
			id = CUSTOM_NUMFMT_BASE + customFormats.size
			customFormats.set(code, id)
		}
		return id
	}

	const internFont = (font: FontStyle | undefined): number => {
		if (font === undefined) return 0
		const key = keyOf(font)
		let index = fontIndex.get(key)
		if (index === undefined) {
			index = fonts.length
			fonts.push(fontXml(font))
			fontIndex.set(key, index)
		}
		if (colorUsesTheme(font.color)) themeUsed = true
		return index
	}

	const internFill = (fill: FillStyle | undefined): number => {
		if (fill === undefined) return 0
		const key = keyOf(fill)
		let index = fillIndex.get(key)
		if (index === undefined) {
			index = fills.length
			fills.push(fillXml(fill))
			fillIndex.set(key, index)
		}
		if (colorUsesTheme(fill.fgColor) || colorUsesTheme(fill.bgColor)) themeUsed = true
		return index
	}

	const internBorder = (border: BorderStyle | undefined): number => {
		if (border === undefined) return 0
		const key = keyOf(border)
		let index = borderIndex.get(key)
		if (index === undefined) {
			index = borders.length
			borders.push(borderXml(border))
			borderIndex.set(key, index)
		}
		for (const edge of [border.top, border.right, border.bottom, border.left]) {
			if (edge !== undefined && colorUsesTheme(edge.color)) themeUsed = true
		}
		return index
	}

	function xfIndexFor(style: CellStyle | undefined, isDate: boolean, ref: string): number {
		let fontId = 0
		let fillId = 0
		let borderId = 0
		let alignment: Alignment | undefined
		let numFmtCode: string | undefined
		if (style !== undefined) {
			if (!isPlainObject(style)) invalid(ref, "style must be an object")
			checkKeys(ref, "style", style as Record<string, unknown>, [
				"font",
				"fill",
				"border",
				"alignment",
				"numberFormat",
			])
			// Single-read each component (see the validator note above) before validating it.
			const rawCode = style.numberFormat
			if (rawCode !== undefined) {
				// A format CODE string — what Excel's Custom dialog shows and numberFormat(ref)
				// returns. Ids are file-internal and never part of the API.
				if (typeof rawCode !== "string" || rawCode.length === 0) {
					invalid(ref, "style.numberFormat must be a non-empty format code string")
				}
				if (!isXmlSafe(rawCode)) {
					invalid(
						ref,
						"style.numberFormat contains a character not allowed in XML (a control character or lone surrogate)",
					)
				}
				numFmtCode = rawCode
			}
			const font = style.font
			fontId = internFont(font === undefined ? undefined : validateFont(ref, font))
			const fill = style.fill
			fillId = internFill(fill === undefined ? undefined : validateFill(ref, fill))
			const border = style.border
			borderId = internBorder(border === undefined ? undefined : validateBorder(ref, border))
			const align = style.alignment
			alignment = align === undefined ? undefined : validateAlignment(ref, align)
		}
		// A user code wins even on a Date — the implicit date format (id 14) applies only when the
		// caller didn't choose one. 'General' reverse-maps to id 0, i.e. no format.
		let numFmtId = 0
		if (numFmtCode !== undefined) numFmtId = numFmtIdFor(numFmtCode)
		else if (isDate) numFmtId = DATE_NUMFMT_ID

		const key = `${numFmtId}/${fontId}/${fillId}/${borderId}/${keyOf(alignment)}`
		let index = xfIndex.get(key)
		if (index === undefined) {
			index = xfs.length
			xfs.push({ numFmtId, fontId, fillId, borderId, alignment })
			xfIndex.set(key, index)
		}
		return index
	}

	function xfXml(xf: XfRecord): string {
		// Attribute + apply-flag layout matches what the pre-F4.2 writer emitted, so the bare-value
		// stylesheet is byte-identical. An apply flag appears only for a non-default component.
		let attrs = `numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"`
		if (xf.numFmtId !== 0) attrs += ' applyNumberFormat="1"'
		if (xf.fontId !== 0) attrs += ' applyFont="1"'
		if (xf.fillId >= 2) attrs += ' applyFill="1"'
		if (xf.borderId !== 0) attrs += ' applyBorder="1"'
		if (xf.alignment !== undefined) {
			return `<xf ${attrs} applyAlignment="1">${alignmentXml(xf.alignment)}</xf>`
		}
		return `<xf ${attrs}/>`
	}

	function stylesXml(): string {
		// Schema order puts <numFmts> FIRST, before <fonts>. Only custom codes are declared —
		// built-in ids are implicit. formatCode is attribute-escaped: real codes carry quotes
		// (e.g. "kg" 0.0).
		const numFmts =
			customFormats.size === 0
				? ""
				: `<numFmts count="${customFormats.size}">${[...customFormats]
						.map(
							([code, id]) =>
								`<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`,
						)
						.join("")}</numFmts>`
		return (
			`${XML_DECL}\n<styleSheet xmlns="${NS_MAIN}">${numFmts}` +
			`<fonts count="${fonts.length}">${fonts.join("")}</fonts>` +
			`<fills count="${fills.length}">${fills.join("")}</fills>` +
			`<borders count="${borders.length}">${borders.join("")}</borders>` +
			'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
			`<cellXfs count="${xfs.length}">${xfs.map(xfXml).join("")}</cellXfs>` +
			'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
		)
	}

	return {
		xfIndexFor,
		needed: () => xfs.length > 1,
		usesTheme: () => themeUsed,
		stylesXml,
	}
}
