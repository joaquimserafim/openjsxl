import { XlsxError } from "../errors";

// XML text serialization for the writer — the encode side of the reader's entity decoding.
//
// OOXML content is UTF-8. Markup characters must be escaped, and XML 1.0 §2.11 normalizes literal
// CR/CRLF in element text before parsing, so a data CR must be emitted as a character reference.
// `&` is replaced first so the `&` introduced by later replacements is not re-escaped.

export function escapeText(s: string): string {
	return s
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/\r/g, "&#13;");
}

export function escapeAttr(s: string): string {
	// Beyond quote-escaping: literal tab/LF/CR inside an ATTRIBUTE value are replaced with spaces
	// by conforming parsers (XML 1.0 §3.3.3 attribute-value normalization) — the same bytes would
	// mean a different string to Excel than to our non-normalizing reader. Character references
	// are exempt from normalization, so emitting &#9;/&#10;/&#13; keeps the value verbatim for
	// every consumer. Element-text CR uses the separate §2.11 handling in escapeText.
	return escapeText(s).replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;");
}

// XML safety lives in utils (shared with the reader's style model, which degrades unsafe
// producer strings so the bridge only ever carries writable values). Re-exported here so the
// writer keeps one import site for its serialization helpers.
export { isXmlSafe } from "../utils";

/**
 * A direct instance of `Object` (or a null-prototype object) — not an array, `Date`, class instance,
 * or otherwise exotic value. This checks shape; callers must still use ownValue to capture each
 * allowed field once and validate it. Shared by sheet-level and workbook-level validators.
 */
export function isPlainRecord(value: unknown): value is Record<string, unknown> {
	if (typeof value !== "object" || value === null) return false;
	const proto = Object.getPrototypeOf(value);
	return proto === null || proto === Object.prototype;
}

/**
 * Capture an own model field once; inherited values are absent, including on sparse arrays.
 * Use Object.hasOwn separately when the schema distinguishes absence from an own undefined value.
 * Retains the original receiver for getters and supports null-prototype records.
 */
export function ownValue<T extends object, K extends keyof T>(record: T, key: K): T[K] | undefined {
	return Object.hasOwn(record, key) ? record[key] : undefined;
}

// The deepest prototype chain the check walks. Real class hierarchies are a few links deep; a Proxy
// whose getPrototypeOf trap mints a new prototype on every call would otherwise never end the walk.
export const MAX_PROTOTYPE_DEPTH = 1000;

/**
 * Reject a record whose prototype chain supplies one of the model fields in `keys` (F11.34). The
 * writer reads only own properties, so a field that a class or `Object.create` prototype supplies,
 * such as a prototype getter, would otherwise be dropped without an error. A key the record owns
 * itself is fine even when a prototype also has it: the own value is the one written. The walk
 * starts at the record's prototype and stops at the first `Object.prototype` or `null`: a plain or
 * null-prototype record costs one prototype read, and a key polluted onto `Object.prototype` stays
 * ignored. It tests keys with `Object.hasOwn` and reads no value, so no getter runs. A chain longer
 * than {@link MAX_PROTOTYPE_DEPTH} is rejected too. A Proxy can hide a field from the check, by
 * serving it from its `get` trap alone or by reporting a plain prototype from its `getPrototypeOf`
 * trap; that field is not written. Model fields must be own properties.
 *
 * @example rejectPrototypeFields({ value: 5 }, ["value", "style"], "cell A1") // returns
 * @example rejectPrototypeFields(new (class { value = 5 })(), ["value", "style"], "cell A1") // returns
 * @example rejectPrototypeFields(new (class { value = 5; get style() { return {}; } })(), ["value", "style"], "cell A1") // throws XlsxError invalid-input: cell A1: "style" comes from a prototype; set it as an own property
 * @example rejectPrototypeFields(endlessProxyChain, ["value", "style"], "cell A1") // throws XlsxError invalid-input: cell A1: its prototype chain is longer than 1000; use a plain object
 */
export function rejectPrototypeFields(record: object, keys: readonly string[], what: string): void {
	let proto: unknown = Object.getPrototypeOf(record);
	for (
		let depth = 1;
		(typeof proto === "object" || typeof proto === "function") &&
		proto !== null &&
		proto !== Object.prototype;
		depth++
	) {
		if (depth > MAX_PROTOTYPE_DEPTH) {
			throw new XlsxError(
				"invalid-input",
				`${what}: its prototype chain is longer than ${MAX_PROTOTYPE_DEPTH}; use a plain object`,
			);
		}
		for (const key of keys) {
			if (Object.hasOwn(proto, key) && !Object.hasOwn(record, key)) {
				throw new XlsxError(
					"invalid-input",
					`${what}: "${key}" comes from a prototype; set it as an own property`,
				);
			}
		}
		proto = Object.getPrototypeOf(proto);
	}
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
