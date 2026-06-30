// XML-name helpers shared across the OOXML parsing layers.

// Strip a namespace prefix from a qualified name: `r:id` -> `id`, `si` -> `si`. The
// tokenizer treats prefixes as part of the literal name (it does no namespace
// resolution), so consumers that only care about the local name normalise here. Excel
// almost always uses the default (unprefixed) namespace, but other producers may not.
export function localName(name: string): string {
	const colon = name.indexOf(':')
	return colon === -1 ? name : name.slice(colon + 1)
}
