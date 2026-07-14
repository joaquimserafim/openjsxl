// Decode the XML character entities that appear in OOXML text: the five predefined
// entities plus decimal and hexadecimal numeric character references. The tokenizer
// uses this for both element text and attribute values.
//
// Per the XML spec, hex references use a lowercase `x` (`&#x41;`); anything that does
// not match a known form is left untouched so malformed input round-trips unchanged.
//
// The numeric-reference digit counts are BOUNDED (≤ 6 hex / ≤ 7 decimal — enough for the
// largest scalar U+10FFFF): a longer run, even one that is only leading-zero padding, is left
// literal. This keeps the longest decodable reference (`&#x10FFFF;` / `&#1114111;` = 10 chars)
// within the streaming tokenizer's split-entity window (xml/stream.ts MAX_ENTITY_TAIL), so a
// reference straddling a chunk boundary decodes identically to a one-shot parse — without the
// bound, a padded reference longer than the window would be held-then-dropped by the stream but
// decoded whole by one-shot (the F9.7 review's split-charref divergence).
const ENTITY_PATTERN = /&(#x[0-9a-fA-F]{1,6}|#[0-9]{1,7}|amp|lt|gt|quot|apos);/g;

export function decodeXmlEntities(input: string): string {
	if (!input.includes("&")) return input;
	return input.replace(ENTITY_PATTERN, (match, body: string) => {
		switch (body) {
			case "amp":
				return "&";
			case "lt":
				return "<";
			case "gt":
				return ">";
			case "quot":
				return '"';
			case "apos":
				return "'";
			default: {
				// Numeric reference: body is "#NNN" (decimal) or "#xHHH" (hex).
				const code =
					body[1] === "x"
						? Number.parseInt(body.slice(2), 16)
						: Number.parseInt(body.slice(1), 10);
				// Reject out-of-range values and lone surrogates (U+D800–U+DFFF): both would
				// produce an ill-formed scalar, so leave the reference as literal text instead.
				const isScalar =
					code >= 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff);
				return isScalar ? String.fromCodePoint(code) : match;
			}
		}
	});
}
