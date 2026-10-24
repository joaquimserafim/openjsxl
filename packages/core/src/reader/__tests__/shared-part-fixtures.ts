import { resolveTarget } from "../../ooxml";
import { writeXlsx } from "../../writer";
import { writeZip } from "../../writer/zip";
import { openZip } from "../../zip";
import { directoryOf, relsPathFor } from "../opc-path";

// F11.95 fixtures: workbooks whose relationships name one package part many times.

/** The relationship type of an OPC table part. */
export const TABLE_REL =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

/**
 * A copy of the archive `bytes` in which each part named in `edits` holds its edit of the part's
 * text. An edited part must exist and the edit must change it, so a fixture cannot silently stop
 * exercising its case.
 *
 * @example await rewriteParts(xlsx, { "xl/_rels/workbook.xml.rels": (t) => t.replace("sheet3.xml", "sheet1.xml") }) // S3 now names sheet1.xml
 * @example await rewriteParts(xlsx, { "missing.xml": (t) => t }) // throws Error: missing.xml is not in the archive
 * @example await rewriteParts(xlsx, { "xl/workbook.xml": (t) => t }) // throws Error: the edit left xl/workbook.xml unchanged
 */
export async function rewriteParts(
	bytes: Uint8Array,
	edits: Readonly<Record<string, (text: string) => string>>,
): Promise<Uint8Array> {
	const zip = openZip(bytes);
	for (const name of Object.keys(edits)) {
		if (!zip.has(name)) throw new Error(`${name} is not in the archive`);
	}
	const parts: { name: string; data: Uint8Array }[] = [];
	for (const name of zip.entries.keys()) {
		const data = await zip.read(name);
		const edit = Object.hasOwn(edits, name) ? edits[name] : undefined;
		if (edit === undefined) {
			parts.push({ name, data });
			continue;
		}
		const text = new TextDecoder().decode(data);
		const edited = edit(text);
		if (edited === text) throw new Error(`the edit left ${name} unchanged`);
		parts.push({ name, data: new TextEncoder().encode(edited) });
	}
	return writeZip(parts);
}

/**
 * Three sheets: S1 (a header row, a table "Sales" on A1:B2 and a comment on A1), S2 (a comment on A1)
 * and S3, with one defined name scoped to no sheet, one to S1 and one to S3.
 *
 * @example (await openXlsx(await threeSheetWorkbook())).sheets.map((s) => s.name) // ["S1", "S2", "S3"]
 */
export function threeSheetWorkbook(): Promise<Uint8Array> {
	return writeXlsx({
		sheets: [
			{
				name: "S1",
				rows: [
					["a", "b"],
					[1, 2],
				],
				comments: [{ ref: "A1", text: "c1" }],
				tables: [
					{ name: "Sales", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false },
				],
			},
			{ name: "S2", rows: [[3]], comments: [{ ref: "A1", text: "c2" }] },
			{ name: "S3", rows: [[4]] },
		],
		definedNames: [
			{ name: "Everywhere", refersTo: "S1!$A$1" },
			{ name: "OnFirst", refersTo: "S1!$A$1", localSheetId: 0 },
			{ name: "OnThird", refersTo: "S3!$A$1", localSheetId: 2 },
		],
	});
}

/** A sheet an xlsx package declares, with the table and comments parts its relationships name. */
export interface DeclaredSheet {
	readonly path: string;
	readonly tables: readonly string[];
	readonly comments: string | undefined;
	/** `<comment>` elements in the comments part, 0 without one. */
	readonly commentCount: number;
}

/**
 * The sheets an xlsx package declares whose part exists, in tab order, with the existing table and
 * comments parts each one's relationships name. It reads the raw XML with regular expressions, apart
 * from the reader, so the corpus pins can check what the reader loads.
 *
 * @example (await declaredSheets(await threeSheetWorkbook())).map((s) => s.path) // ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]
 * @example await declaredSheets(new Uint8Array(4)) // throws XlsxError not-a-zip
 */
export async function declaredSheets(bytes: Uint8Array): Promise<DeclaredSheet[]> {
	const zip = openZip(bytes);
	const text = async (path: string) => new TextDecoder().decode(await zip.read(path));
	const office = relationships(await text("_rels/.rels")).find((r) =>
		r.type.endsWith("/officeDocument"),
	);
	if (office === undefined) return [];
	const workbookPath = resolveTarget("", office.target);
	const workbookRels = relationships(await text(relsPathFor(workbookPath)));
	const sheets: DeclaredSheet[] = [];
	for (const tag of (await text(workbookPath)).match(/<(?:\w+:)?sheet\b[^>]*>/g) ?? []) {
		const id = /\s\w+:id="([^"]*)"/.exec(tag)?.[1];
		const rel = workbookRels.find((r) => r.id === id && !r.external);
		if (rel === undefined) continue;
		const path = resolveTarget(directoryOf(workbookPath), rel.target);
		if (!zip.has(path)) continue;
		const sheetRels = zip.has(relsPathFor(path))
			? relationships(await text(relsPathFor(path)))
			: [];
		const linked = (kind: string) =>
			sheetRels
				.filter((r) => r.type.endsWith(`/${kind}`) && !r.external)
				.map((r) => resolveTarget(directoryOf(path), r.target))
				.filter((part) => zip.has(part));
		const comments = linked("comments")[0];
		const commentCount =
			comments === undefined ? 0 : ((await text(comments)).match(/<comment\b/g) ?? []).length;
		sheets.push({ path, tables: linked("table"), comments, commentCount });
	}
	return sheets;
}

/**
 * The `<Relationship>` elements of a rels part: id, type, target and whether the target is external.
 *
 * @example relationships('<Relationship Id="rId1" Type="t/worksheet" Target="a.xml"/>') // [{ id: "rId1", type: "t/worksheet", target: "a.xml", external: false }]
 * @example relationships("<Relationships/>") // []
 */
export function relationships(xml: string): {
	readonly id: string;
	readonly type: string;
	readonly target: string;
	readonly external: boolean;
}[] {
	return (xml.match(/<Relationship\b[^>]*>/g) ?? []).map((tag) => {
		const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? "";
		return {
			id: attr("Id"),
			type: attr("Type"),
			target: attr("Target"),
			external: attr("TargetMode") === "External",
		};
	});
}
