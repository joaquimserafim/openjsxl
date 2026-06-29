// Decode the XML character entities that appear in OOXML text: the five predefined
// entities plus decimal and hexadecimal numeric character references. The tokenizer
// uses this for both element text and attribute values.
//
// Per the XML spec, hex references use a lowercase `x` (`&#x41;`); anything that does
// not match a known form is left untouched so malformed input round-trips unchanged.

const ENTITY_PATTERN = /&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g

export function decodeXmlEntities(input: string): string {
	if (!input.includes('&')) return input
	return input.replace(ENTITY_PATTERN, (match, body: string) => {
		switch (body) {
			case 'amp':
				return '&'
			case 'lt':
				return '<'
			case 'gt':
				return '>'
			case 'quot':
				return '"'
			case 'apos':
				return "'"
			default: {
				// Numeric reference: body is "#NNN" (decimal) or "#xHHH" (hex).
				const code =
					body[1] === 'x'
						? Number.parseInt(body.slice(2), 16)
						: Number.parseInt(body.slice(1), 10)
				return code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match
			}
		}
	})
}
