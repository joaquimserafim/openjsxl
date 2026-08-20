import { existsSync } from "node:fs";
import { loadFixture, loadLocalFixture, localFixturePath } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";
import { parseComments } from "../worksheet";

// Comments (F2.3). They live in a separate part (xl/commentsN.xml) linked from the worksheet
// rels: an <authors> list plus a <commentList> whose <comment ref authorId> holds rich text.
// parseComments concatenates the <t> runs and resolves the author. The parser is covered by
// units here; real Excel output is verified against a local-only Apache-2.0 POI file (skipped
// when absent — see packages/fixtures/local).

describe("parseComments — units", () => {
	it("resolves the author and concatenates rich-text runs", () => {
		const xml =
			"<comments><authors><author>Ada</author></authors>" +
			'<commentList><comment ref="B2" authorId="0"><text>' +
			'<r><rPr><b/></rPr><t>bold </t></r><r><t xml:space="preserve">and plain</t></r>' +
			"</text></comment></commentList></comments>";
		expect(parseComments(xml)).toEqual([{ ref: "B2", author: "Ada", text: "bold and plain" }]);
	});

	it("reads plain <t> comments and a second author", () => {
		const xml =
			"<comments><authors><author>A</author><author>B</author></authors><commentList>" +
			'<comment ref="A1" authorId="0"><text><t>hi</t></text></comment>' +
			'<comment ref="C3" authorId="1"><text><t>yo</t></text></comment>' +
			"</commentList></comments>";
		expect(parseComments(xml)).toEqual([
			{ ref: "A1", author: "A", text: "hi" },
			{ ref: "C3", author: "B", text: "yo" },
		]);
	});

	it("decodes each complete legacy author as ST_Xstring before resolving authorId", () => {
		const xml =
			"<comments><authors>" +
			"<author>_x0041_</author>" + // raw escape → A
			"<author>_x005F_x0041_</author>" + // protected literal
			"<author>_x005f_x0041_</author>" + // Excel Online's lowercase save-back form
			"<author>_x00<![CDATA[41_]]></author>" + // one escape split across text tokens
			"<author>_x005F_<![CDATA[x0041_]]></author>" + // protected form must not decode twice
			"<author/>" +
			"</authors><commentList>" +
			'<comment ref="A1" authorId="0"><text><t>a</t></text></comment>' +
			'<comment ref="A2" authorId="1"><text><t>b</t></text></comment>' +
			'<comment ref="A3" authorId="2"><text><t>c</t></text></comment>' +
			'<comment ref="A4" authorId="3"><text><t>d</t></text></comment>' +
			'<comment ref="A5" authorId="4"><text><t>e</t></text></comment>' +
			'<comment ref="A6" authorId="5"><text><t>f</t></text></comment>' +
			"</commentList></comments>";
		expect(parseComments(xml)).toEqual([
			{ ref: "A1", author: "A", text: "a" },
			{ ref: "A2", author: "_x0041_", text: "b" },
			{ ref: "A3", author: "_x0041_", text: "c" },
			{ ref: "A4", author: "A", text: "d" },
			{ ref: "A5", author: "_x0041_", text: "e" },
			{ ref: "A6", text: "f" },
		]);
	});

	it("omits the author when the authorId resolves to nothing", () => {
		const xml =
			"<comments><authors><author>A</author></authors>" +
			'<commentList><comment ref="A1" authorId="9"><text><t>x</t></text></comment></commentList></comments>';
		expect(parseComments(xml)).toEqual([{ ref: "A1", text: "x" }]);
	});

	it("omits the author when a comment has no authorId (no default to author 0)", () => {
		const xml =
			"<comments><authors><author>A</author></authors>" +
			'<commentList><comment ref="A1"><text><t>x</t></text></comment></commentList></comments>';
		expect(parseComments(xml)).toEqual([{ ref: "A1", text: "x" }]);
	});

	it("does not fabricate author 0 for empty or non-integer authorId values", () => {
		const xml =
			"<comments><authors><author>A</author></authors><commentList>" +
			'<comment ref="A1" authorId=""><text><t>empty</t></text></comment>' +
			'<comment ref="A2" authorId="x"><text><t>text</t></text></comment>' +
			'<comment ref="A3" authorId="1.5"><text><t>fraction</t></text></comment>' +
			"</commentList></comments>";
		expect(parseComments(xml)).toEqual([
			{ ref: "A1", text: "empty" },
			{ ref: "A2", text: "text" },
			{ ref: "A3", text: "fraction" },
		]);
	});

	it("tolerates a namespace prefix and an empty comment", () => {
		const xml =
			"<x:comments><x:authors><x:author>A</x:author></x:authors>" +
			'<x:commentList><x:comment ref="A1" authorId="0"><x:text/></x:comment></x:commentList></x:comments>';
		expect(parseComments(xml)).toEqual([{ ref: "A1", author: "A", text: "" }]);
	});

	it("is empty when there are no comments", () => {
		expect(parseComments("<comments><authors/><commentList/></comments>")).toEqual([]);
	});

	it("decodes ST_Xstring escapes in comment text, per <t> run (F9.6)", () => {
		const xml =
			"<comments><authors><author>A</author></authors><commentList>" +
			'<comment ref="A1" authorId="0"><text><t>a_x000B_b</t></text></comment>' +
			'<comment ref="B1" authorId="0"><text><t>_x005F_x0041_</t></text></comment>' + // protected literal
			'<comment ref="C1" authorId="0"><text><r><t>_x00</t></r><r><t>41_</t></r></text></comment>' + // no straddle
			"</commentList></comments>";
		expect(parseComments(xml).map((c) => c.text)).toEqual(["a\x0Bb", "_x0041_", "_x0041_"]);
	});
});

describe("Worksheet.comments — no comments part", () => {
	it("is empty for a sheet without comments", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"));
		expect(wb.sheet("Sheet1").comments).toEqual([]);
	});
});

// Real-producer e2e (openpyxl 3.1.5 — see fixtures/data/README.md). Committed to the corpus so it
// always runs, unlike the local-only POI file above. Covers two resolved authors plus an
// author-less comment (openpyxl writes <author/>, which the reader omits rather than fabricating).
describe("Worksheet.comments — openpyxl-authored fixture (e2e)", () => {
	it("reads resolved authors and omits the author-less one", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-comments.xlsx"));
		expect(wb.sheet("Notes").comments).toEqual([
			{ ref: "B2", author: "Ada", text: "check this figure" },
			{ ref: "C3", author: "Grace", text: "EMEA only" },
			{ ref: "D4", text: "no attribution" }, // <author/> → author omitted
		]);
		expect(wb.sheet("Plain").comments).toEqual([]);
	});
});

// Real Excel comment output — Apache-2.0 POI fixtures kept local-only (git/npm-ignored).
// Runs when present; skipped on a fresh clone or in CI.
const poi = "SimpleWithComments.xlsx";

describe.skipIf(!existsSync(localFixturePath(poi)))(
	"Worksheet.comments — real Excel (local)",
	() => {
		it("reads authored, multi-line comments from genuine Excel output", async () => {
			const wb = await openXlsx(await loadLocalFixture(poi));
			expect(wb.sheet("Sheet1").comments).toEqual([
				{ ref: "B1", author: "Yegor Kozlov", text: "Yegor Kozlov:\r\nfirst cell" },
				{ ref: "B2", author: "Yegor Kozlov", text: "Yegor Kozlov:\r\nsecond cell" },
				{ ref: "B3", author: "Yegor Kozlov", text: "Yegor Kozlov:\r\nthird cell" },
			]);
		});
	},
);
