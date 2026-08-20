// OPC relationship parts live beside their source part in a `_rels` directory.

/** Return the containing OPC directory: `xl/workbook.xml` → `xl`. */
export function directoryOf(path: string): string {
	const slash = path.lastIndexOf("/");
	return slash === -1 ? "" : path.slice(0, slash);
}

/** Return a part's relationship path: `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. */
export function relsPathFor(path: string): string {
	const slash = path.lastIndexOf("/");
	const dir = slash === -1 ? "" : path.slice(0, slash);
	const file = slash === -1 ? path : path.slice(slash + 1);
	return dir === "" ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`;
}
