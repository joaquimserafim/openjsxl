// XML text serialization for the writer — the encode side of the reader's entity decoding.
//
// OOXML content is UTF-8, so the only characters that MUST be escaped are the ones that would
// otherwise be read as markup: `&`, `<`, `>` in text, plus `"` inside a double-quoted attribute.
// `&` is replaced first so the `&` we introduce for the others isn't re-escaped. We do not emit
// numeric character references for anything else — the bytes go out as UTF-8.

export function escapeText(s: string): string {
	return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function escapeAttr(s: string): string {
	// Beyond quote-escaping: literal tab/LF/CR inside an ATTRIBUTE value are replaced with spaces
	// by conforming parsers (XML 1.0 §3.3.3 attribute-value normalization) — the same bytes would
	// mean a different string to Excel than to our non-normalizing reader. Character references
	// are exempt from normalization, so emitting &#9;/&#10;/&#13; keeps the value verbatim for
	// every consumer. Element TEXT is not normalized, so escapeText stays as-is.
	return escapeText(s)
		.replace(/"/g, "&quot;")
		.replace(/\t/g, "&#9;")
		.replace(/\n/g, "&#10;")
		.replace(/\r/g, "&#13;");
}

// XML safety lives in utils (shared with the reader's style model, which degrades unsafe
// producer strings so the bridge only ever carries writable values). Re-exported here so the
// writer keeps one import site for its serialization helpers.
export { isXmlSafe } from "../utils";

/**
 * A direct instance of `Object` (or a null-prototype object) — not an array, `Date`, class instance,
 * or otherwise exotic value. Writer input validation rejects anything else BEFORE reading properties,
 * so a hostile prototype/getter can't smuggle values past validation (single-read TOCTOU). Shared by
 * the sheet-level and workbook-level validators.
 */
export function isPlainRecord(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false;
	const proto = Object.getPrototypeOf(value);
	return proto === null || proto === Object.prototype;
}

// A string whose meaningful leading/trailing whitespace would be stripped by an XML reader needs
// `xml:space="preserve"` on its element, or Excel drops those spaces on load. We flag it whenever
// trimming would change the string (covers leading/trailing spaces, tabs, and newlines).
export function needsPreserve(s: string): boolean {
	return s !== s.trim();
}

/** The ` xml:space="preserve"` attribute (with leading space) when `s` needs it, else `''`. */
export function preserveAttr(s: string): string {
	return needsPreserve(s) ? ' xml:space="preserve"' : "";
}
