// ST_Xstring — the OOXML escape convention for string CONTENT (shared strings, inline strings,
// cached formula strings, comment text; ECMA-376 §22.9.2.19 / MS-OI29500 §2.1.1774). Characters
// XML 1.0 cannot carry faithfully (C0 controls except tab/LF — including CR, which XML parsers
// silently normalize to LF — U+FFFE/U+FFFF, lone surrogates) are stored as `_xHHHH_` — one UTF-16
// code unit in hex — and a LITERAL substring that happens to look like an escape is protected by
// escaping its own leading underscore: the data `_x0041_` is stored as `_x005F_x0041_`. Excel and
// openpyxl both apply this on load/save, so a writer that emits the pattern verbatim silently
// corrupts such data for every other consumer (F9.6).
//
// The pair is single-sourced here for the SHARED-BOUNDS invariant: whatever character the reader
// can now hand back (a decoded control char, even a lone surrogate a producer escaped), the
// writer re-encodes rather than rejects — `decodeXstring(encodeXstring(s)) === s` for EVERY
// string, and `encodeXstring` is the IDENTITY on strings with nothing to escape, which is what
// keeps clean files byte-identical.

// One UTF-16 code unit of a hex quad: 0-9, A-F, a-f. `charCodeAt` past the end is NaN — false.
function isHexUnit(c: number): boolean {
	return (c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x46) || (c >= 0x61 && c <= 0x66);
}

// True when the 7 units starting at `i` spell `_xHHHH_` (the `x` lowercase, as Excel writes and
// openpyxl matches; hex either case).
function isEscapeShape(s: string, i: number): boolean {
	return (
		s.charCodeAt(i) === 0x5f &&
		s.charCodeAt(i + 1) === 0x78 &&
		isHexUnit(s.charCodeAt(i + 2)) &&
		isHexUnit(s.charCodeAt(i + 3)) &&
		isHexUnit(s.charCodeAt(i + 4)) &&
		isHexUnit(s.charCodeAt(i + 5)) &&
		s.charCodeAt(i + 6) === 0x5f
	);
}

/**
 * Decode `_xHHHH_` escapes in stored string content — a single left-to-right, NON-overlapping,
 * non-reentrant pass (a decoded `_x005F_` yields a literal `_` that is never re-examined, so
 * `_x005F_x0041_` decodes to the data `_x0041_`, exactly as Excel and openpyxl's `re.sub` do).
 * Text with no escape comes back unchanged (identity).
 */
export function decodeXstring(s: string): string {
	let i = s.indexOf("_x");
	if (i === -1) return s; // fast path: no candidate anywhere
	let out = "";
	let start = 0; // beginning of the yet-uncopied literal run
	while (i !== -1) {
		if (isEscapeShape(s, i)) {
			out += s.slice(start, i);
			out += String.fromCharCode(Number.parseInt(s.slice(i + 2, i + 6), 16));
			start = i + 7;
			i = s.indexOf("_x", start);
		} else {
			i = s.indexOf("_x", i + 1);
		}
	}
	return out + s.slice(start);
}

// Whether the unit at `p` is stored as an `_xHHHH_` escape: the set `isXmlSafe` rejects (C0
// controls except tab/LF, U+FFFE/U+FFFF, lone surrogates) PLUS carriage return — XML 1.0 §2.11
// line-ending normalization turns a raw CR into LF in every conforming parser, so Excel stores
// CR as `_x000D_` and so do we (adversarial-review fix). A high surrogate followed by its low
// half is a well-formed pair, not escaped; a LOW surrogate examined here is always lone (the
// encode loop consumes pairs whole, and the collision lookahead only lands after a hex digit).
function unitNeedsEscape(s: string, p: number): boolean {
	const c = s.charCodeAt(p);
	if (Number.isNaN(c)) return false; // past the end
	if (c < 0x20) return c !== 0x09 && c !== 0x0a;
	if (c === 0xfffe || c === 0xffff) return true;
	if (c >= 0xd800 && c <= 0xdbff) {
		const next = s.charCodeAt(p + 1);
		return !(next >= 0xdc00 && next <= 0xdfff);
	}
	return c >= 0xdc00 && c <= 0xdfff;
}

/**
 * Encode string content for storage: each XML-unrepresentable UTF-16 code unit (see
 * {@link unitNeedsEscape} — `isXmlSafe`'s rejection set plus CR) becomes `_xHHHH_`, and a
 * literal underscore that a decoder would see as starting an escape is itself escaped as
 * `_x005F_`. That collision check must consider the OUTPUT, not just the input: a literal
 * `_xHHHH` prefix followed by a unit that gets escaped would have its shape COMPLETED by the
 * emitted escape's leading underscore (`"_x0041" + "\x01"` → `"_x0041_x0001_"` reads back as
 * `"Ax0001_"` — the adversarial-review HIGH), so that underscore is protected too. Well-formed
 * astral pairs (emoji) pass through untouched. A string with nothing to escape is returned
 * as-is — the identity that preserves byte-identity for clean input.
 */
export function encodeXstring(s: string): string {
	let out = "";
	let start = 0; // beginning of the yet-uncopied literal run
	let i = 0;
	while (i < s.length) {
		const c = s.charCodeAt(i);
		if (c >= 0xd800 && c <= 0xdbff) {
			const next = s.charCodeAt(i + 1); // NaN past the end — fails the range test
			if (next >= 0xdc00 && next <= 0xdfff) {
				i += 2; // a well-formed astral pair is XML-safe — keep both units
				continue;
			}
		}
		let esc: boolean;
		if (c === 0x5f) {
			// Protect a literal `_` that would START an escape shape in the OUTPUT: either the
			// input already spells `_xHHHH_`, or it spells `_xHHHH` and the next unit will
			// itself be escaped — the emitted escape's `_` would close the shape.
			esc =
				s.charCodeAt(i + 1) === 0x78 &&
				isHexUnit(s.charCodeAt(i + 2)) &&
				isHexUnit(s.charCodeAt(i + 3)) &&
				isHexUnit(s.charCodeAt(i + 4)) &&
				isHexUnit(s.charCodeAt(i + 5)) &&
				(s.charCodeAt(i + 6) === 0x5f || unitNeedsEscape(s, i + 6));
		} else {
			esc = unitNeedsEscape(s, i);
		}
		if (esc) {
			out += s.slice(start, i);
			out += `_x${c.toString(16).toUpperCase().padStart(4, "0")}_`;
			start = i + 1;
		}
		i++;
	}
	return start === 0 ? s : out + s.slice(start);
}
