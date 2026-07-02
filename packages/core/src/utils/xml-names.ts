// XML-name helpers shared across the OOXML parsing layers.

// Strip a namespace prefix from a qualified name: `r:id` -> `id`, `si` -> `si`. The
// tokenizer treats prefixes as part of the literal name (it does no namespace
// resolution), so consumers that only care about the local name normalise here. Excel
// almost always uses the default (unprefixed) namespace, but other producers may not.
export function localName(name: string): string {
	const colon = name.indexOf(":")
	return colon === -1 ? name : name.slice(colon + 1)
}

// The relationship id on an element is conventionally `r:id`, but the `r` prefix is only
// bound by convention — fall back to any attribute whose local name is `id`. Shared by the
// workbook (<sheet r:id>) and worksheet (<hyperlink r:id>) parsers.
//
// This trusts the prefix/local name rather than resolving the relationships namespace URI —
// the tokenizer does no namespace resolution, by design. It is safe for the elements we call
// it on: <sheet> and <hyperlink> carry no attribute whose local name is `id` other than the
// relationship id, so the fallback cannot mis-match. Reusing it on an element that may carry a
// foreign `*:id` attribute would need revisiting.
export function relationshipId(attrs: Readonly<Record<string, string>>): string | undefined {
	if (attrs["r:id"] !== undefined) return attrs["r:id"]
	for (const key of Object.keys(attrs)) {
		if (localName(key) === "id") return attrs[key]
	}
	return undefined
}
