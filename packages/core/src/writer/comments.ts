import { parseCanonicalCell } from "../ooxml/a1";
import { encodeXstring } from "../ooxml/xstring";
import type { Comment } from "../types";
import { checkKeys, sheetInvalid, shortened } from "./sheet-validation";
import { escapeText, isPlainRecord, isXmlSafe, ownValue, preserveAttr } from "./xml";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const VML_HEAD =
	'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
	'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
	'<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
const VML_TAIL = "</xml>";

/**
 * Build one hidden legacy-note shape from zero-based cell coordinates.
 *
 * Input: index `0`, row `1`, column `1` (cell B2).
 * Output: shape `_x0000_s1025` containing `<x:Row>1</x:Row><x:Column>1</x:Column>`.
 */
export function vmlShape(index: number, row0: number, col0: number): string {
	return (
		`<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" ` +
		`style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144px;height:79px;z-index:${index + 1};visibility:hidden" ` +
		'fillcolor="#ffffe1" o:insetmode="auto">' +
		'<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/>' +
		'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>' +
		'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill>' +
		`<x:Row>${row0}</x:Row><x:Column>${col0}</x:Column></x:ClientData></v:shape>`
	);
}

export interface CommentsParts {
	readonly commentsXml: string;
	readonly vmlXml: string;
}

/**
 * Validate comments and build their OOXML comments and legacy VML parts.
 *
 * Input: `[{ ref: "B2", author: "Ada", text: "Note" }]`.
 * Output: one author/comment entry and one VML shape anchored at row 1, column 1.
 */
export function commentsParts(
	sheetName: string,
	comments: readonly Comment[] | undefined,
): CommentsParts | undefined {
	if (comments === undefined) return undefined;
	if (!Array.isArray(comments)) sheetInvalid(sheetName, "comments must be an array");
	if (comments.length === 0) return undefined;
	const authors: string[] = [];
	const authorIndex = new Map<string, number>();
	const items: {
		readonly ref: string;
		readonly authorId: number;
		readonly text: string;
		readonly row: number;
		readonly col: number;
	}[] = [];
	for (let i = 0; i < comments.length; i++) {
		const raw: unknown = ownValue(comments, i);
		const what = `comments[${i}]`;
		if (!isPlainRecord(raw)) sheetInvalid(sheetName, `${what} must be an object`);
		checkKeys(sheetName, what, raw, ["ref", "author", "text"]);

		const ref = ownValue(raw, "ref");
		if (typeof ref !== "string") sheetInvalid(sheetName, `${what}.ref must be a string`);
		const cell = parseCanonicalCell(ref);
		if (cell === undefined) {
			sheetInvalid(
				sheetName,
				`${what}.ref "${shortened(ref)}" is not a canonical A1 cell within Excel's grid`,
			);
		}

		const rawText = ownValue(raw, "text");
		if (typeof rawText !== "string") sheetInvalid(sheetName, `${what}.text must be a string`);
		const text = encodeXstring(rawText);

		const rawAuthor = ownValue(raw, "author");
		let key: string;
		if (rawAuthor === undefined) key = "";
		else {
			if (typeof rawAuthor !== "string") {
				sheetInvalid(sheetName, `${what}.author must be a string`);
			}
			if (!isXmlSafe(rawAuthor)) {
				sheetInvalid(
					sheetName,
					`${what}.author contains a character not allowed in XML (a control character or lone surrogate)`,
				);
			}
			key = rawAuthor;
		}
		let id = authorIndex.get(key);
		if (id === undefined) {
			id = authors.length;
			authors.push(key);
			authorIndex.set(key, id);
		}
		items.push({ ref, authorId: id, text, row: cell.row, col: cell.col });
	}

	const authorsXml = authors
		.map((author) => `<author>${escapeText(encodeXstring(author))}</author>`)
		.join("");
	const listXml = items
		.map(
			(comment) =>
				`<comment ref="${comment.ref}" authorId="${comment.authorId}" shapeId="0"><text><t${preserveAttr(comment.text)}>${escapeText(comment.text)}</t></text></comment>`,
		)
		.join("");
	const commentsXml = `${XML_DECL}\n<comments xmlns="${NS_MAIN}"><authors>${authorsXml}</authors><commentList>${listXml}</commentList></comments>`;
	const vmlXml = `${VML_HEAD}${items
		.map((comment, index) => vmlShape(index, comment.row - 1, comment.col - 1))
		.join("")}${VML_TAIL}`;
	return { commentsXml, vmlXml };
}
