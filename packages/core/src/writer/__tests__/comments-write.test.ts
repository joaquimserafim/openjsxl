import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { writeXlsx } from "../workbook";

// F5.2 — comments write (legacy VML). A comment renders in Excel only with BOTH the comments part
// (xl/commentsN.xml) and a paired legacy VML drawing; the writer emits both. Everything written
// must re-read through Worksheet.comments verbatim and carry across the bridge, and a comment-free
// sheet must keep its exact pre-F5.2 bytes (the golden pins cover that; the no-op tests re-assert
// it here). The per-sheet rels part is now shared: hyperlinks (F4.6), the comments part, and the
// vmlDrawing part draw non-colliding rIds from one counter.

const decoder = new TextDecoder();
const part = async (bytes: Uint8Array, name: string): Promise<string> =>
	decoder.decode(await openZip(bytes).read(name));

describe("writeXlsx — comments round-trip", () => {
	it("round-trips comments through the reader, omitting the author-less one", async () => {
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "Notes",
						rows: [["v"]],
						comments: [
							{ ref: "B2", author: "Ada", text: "first" },
							{ ref: "C3", author: "Ada", text: "second" }, // same author — dedup is invisible on read
							{ ref: "D4", text: "anon" }, // no author → reader omits it
						],
					},
				],
			}),
		);
		expect(wb.sheet("Notes").comments).toEqual([
			{ ref: "B2", author: "Ada", text: "first" },
			{ ref: "C3", author: "Ada", text: "second" },
			{ ref: "D4", text: "anon" },
		]);
	});

	it("keeps a comment on an otherwise-empty cell (the blank anchor cell is not re-emitted)", async () => {
		// A comment is anchored by ref, not by a cell element — so it survives even when its cell has
		// no value. The empty anchor cell drops like every unstyled empty cell; the comment does not.
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [["v"]],
						comments: [{ ref: "Z9", author: "Ada", text: "blank" }],
					},
				],
			}),
		);
		expect(wb.sheet("S").comments).toEqual([{ ref: "Z9", author: "Ada", text: "blank" }]);
		expect(wb.sheet("S").cell("Z9").type).toBe("empty");
	});
});

describe("writeXlsx — comments part (xl/commentsN.xml)", () => {
	it("builds a first-occurrence-ordered, deduped authors table; author-less shares the empty entry", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["v"]],
					comments: [
						{ ref: "B2", author: "Bob", text: "b" },
						{ ref: "C3", author: "Ada", text: "a" },
						{ ref: "D4", author: "Bob", text: "b2" }, // dup author → reuses id 0
						{ ref: "E5", text: "none" }, // author-less → shared "" author, id 2
					],
				},
			],
		});
		const xml = await part(bytes, "xl/comments1.xml");
		expect(xml).toContain(
			"<authors><author>Bob</author><author>Ada</author><author></author></authors>",
		);
		expect(xml).toContain('<comment ref="B2" authorId="0" shapeId="0">');
		expect(xml).toContain('<comment ref="C3" authorId="1" shapeId="0">');
		expect(xml).toContain('<comment ref="D4" authorId="0" shapeId="0">');
		expect(xml).toContain('<comment ref="E5" authorId="2" shapeId="0">');
	});

	it("escapes comment text and preserves surrounding whitespace", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["v"]],
					comments: [
						{ ref: "A1", author: "P&Q <co>", text: "a & b <c>" },
						{ ref: "A2", text: "  padded  " },
					],
				},
			],
		});
		const xml = await part(bytes, "xl/comments1.xml");
		expect(xml).toContain("<author>P&amp;Q &lt;co&gt;</author>");
		expect(xml).toContain("<text><t>a &amp; b &lt;c&gt;</t></text>");
		expect(xml).toContain('<text><t xml:space="preserve">  padded  </t></text>');
	});
});

describe("writeXlsx — VML legacy drawing", () => {
	it("emits one hidden note shape per comment, positioned by 0-based Row/Column", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["v"]],
					comments: [
						{ ref: "B2", text: "x" },
						{ ref: "D4", text: "y" },
					],
				},
			],
		});
		const vml = await part(bytes, "xl/drawings/vmlDrawing1.vml");
		expect(vml.match(/<v:shape /g)).toHaveLength(2);
		expect(vml).toContain('id="_x0000_s1025"');
		expect(vml).toContain('id="_x0000_s1026"');
		expect(vml).toContain("visibility:hidden");
		// B2 → 0-based row 1, col 1; D4 → row 3, col 3.
		expect(vml).toContain("<x:Row>1</x:Row><x:Column>1</x:Column>");
		expect(vml).toContain("<x:Row>3</x:Row><x:Column>3</x:Column>");
	});
});

describe("writeXlsx — worksheet element + shared rels part", () => {
	it("places <legacyDrawing> after </hyperlinks>, with xmlns:r on the root", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["v"]],
					hyperlinks: [{ ref: "A1", target: "https://example.com" }],
					comments: [{ ref: "B2", text: "x" }],
				},
			],
		});
		const xml = await part(bytes, "xl/worksheets/sheet1.xml");
		expect(xml).toContain(
			'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
		);
		expect(xml).toContain('</hyperlinks><legacyDrawing r:id="rId3"/></worksheet>');
	});

	it("draws non-colliding rIds for hyperlink, comments, and vmlDrawing from one counter", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["v"]],
					hyperlinks: [{ ref: "A1", target: "https://example.com" }],
					comments: [{ ref: "B2", text: "x" }],
				},
			],
		});
		const rels = await part(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
		expect(rels).toContain(
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
		);
		expect(rels).toContain(
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>',
		);
		expect(rels).toContain(
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>',
		);
	});

	it("a comments-only sheet uses rId1/rId2 and declares xmlns:r for the legacyDrawing", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [["v"]], comments: [{ ref: "B2", text: "x" }] }],
		});
		const xml = await part(bytes, "xl/worksheets/sheet1.xml");
		expect(xml).toContain("xmlns:r=");
		expect(xml).not.toContain("<hyperlinks>");
		expect(xml).toContain('</sheetData><legacyDrawing r:id="rId2"/></worksheet>');
		const rels = await part(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
		expect(rels).toContain(
			'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"',
		);
		expect(rels).toContain(
			'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"',
		);
	});
});

describe("writeXlsx — content types & part naming", () => {
	it("registers the vml Default and a per-sheet comments Override", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{ name: "A", rows: [["v"]], comments: [{ ref: "B2", text: "x" }] },
				{ name: "B", rows: [["v"]] }, // no comments
			],
		});
		const ct = await part(bytes, "[Content_Types].xml");
		expect(ct).toContain(
			'<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
		);
		expect(ct).toContain(
			'<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
		);
		expect(ct).not.toContain("/xl/comments2.xml");
		const zip = openZip(bytes);
		expect(zip.has("xl/comments1.xml")).toBe(true);
		expect(zip.has("xl/drawings/vmlDrawing1.vml")).toBe(true);
		expect(zip.has("xl/comments2.xml")).toBe(false);
	});

	it("names parts by sheet index, so a later commented sheet gets commentsN", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{ name: "A", rows: [["v"]] }, // no comments
				{ name: "B", rows: [["v"]], comments: [{ ref: "B2", text: "x" }] },
			],
		});
		const zip = openZip(bytes);
		expect(zip.has("xl/comments2.xml")).toBe(true);
		expect(zip.has("xl/drawings/vmlDrawing2.vml")).toBe(true);
		expect(zip.has("xl/comments1.xml")).toBe(false);
		expect(await part(bytes, "xl/worksheets/_rels/sheet2.xml.rels")).toContain(
			'Target="../comments2.xml"',
		);
	});

	it("a comment-free workbook emits no vml Default and no comment parts", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["v"]] }] });
		expect(await part(bytes, "[Content_Types].xml")).not.toContain('Extension="vml"');
		const zip = openZip(bytes);
		expect(zip.has("xl/comments1.xml")).toBe(false);
		expect(zip.has("xl/drawings/vmlDrawing1.vml")).toBe(false);
		expect(await part(bytes, "xl/worksheets/sheet1.xml")).not.toContain("legacyDrawing");
	});

	it("an empty comments array is a no-op (no parts, no legacyDrawing)", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["v"]], comments: [] }] });
		const zip = openZip(bytes);
		expect(zip.has("xl/comments1.xml")).toBe(false);
		expect(zip.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(false);
		expect(await part(bytes, "xl/worksheets/sheet1.xml")).not.toContain("legacyDrawing");
	});
});

describe("writeXlsx — comment validation", () => {
	const reject = async (comments: unknown, pattern: RegExp): Promise<void> => {
		const err = await writeXlsx({
			// biome-ignore lint/suspicious/noExplicitAny: exercising hostile input the types forbid
			sheets: [{ name: "S", rows: [["v"]], comments: comments as any }],
		}).then(
			() => undefined,
			(e) => e,
		);
		expect(err, String(pattern)).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("invalid-input");
		expect((err as XlsxError).message, String(pattern)).toMatch(pattern);
	};

	it("rejects a non-array comments and a non-object entry", async () => {
		await reject("x", /comments must be an array/);
		await reject([null], /comments\[0\] must be an object/);
		await reject([42], /comments\[0\] must be an object/);
	});

	it("rejects an unknown property", async () => {
		await reject([{ ref: "A1", text: "x", note: "y" }], /unknown property "note"/);
	});

	it("rejects a ref that is not a single canonical cell (incl. ranges and out-of-grid)", async () => {
		for (const ref of ["a1", "A0", "A1:B2", "AAAA1", "XFE1", "A1048577"]) {
			await reject([{ ref, text: "x" }], /not a canonical A1 cell within Excel's grid/);
		}
	});

	it("rejects a missing or non-string text", async () => {
		await reject([{ ref: "A1" }], /text must be a string/);
		await reject([{ ref: "A1", text: 5 }], /text must be a string/);
	});

	it("rejects a non-string author", async () => {
		await reject([{ ref: "A1", text: "x", author: 5 }], /author must be a string/);
	});

	it("stores XML-unsafe text via the ST_Xstring escape; the author still rejects (F9.6)", async () => {
		// Comment TEXT is string content — a control char now stores as _xHHHH_ (the convention
		// Excel decodes) and round-trips. The AUTHOR stays identifier-strict: typed rejection.
		const bytes = await writeXlsx({
			sheets: [
				{ name: "S", rows: [["v"]], comments: [{ ref: "A1", text: "bad\u0001text" }] },
			],
		});
		expect(await part(bytes, "xl/comments1.xml")).toContain("bad_x0001_text");
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").comments).toEqual([{ ref: "A1", text: "bad\u0001text" }]);
		await reject(
			[{ ref: "A1", text: "ok", author: "bad\u0001author" }],
			/author contains a character not allowed in XML/,
		);
	});
});

describe("writeXlsx — TOCTOU single-read", () => {
	it("reads each comment property exactly once (a value-flipping getter cannot diverge validation from emission)", async () => {
		let refReads = 0;
		let textReads = 0;
		let authorReads = 0;
		const comment = {
			get ref() {
				refReads++;
				return "A1";
			},
			get text() {
				textReads++;
				return "hi";
			},
			get author() {
				authorReads++;
				return "Ada";
			},
		};
		await writeXlsx({ sheets: [{ name: "S", rows: [["v"]], comments: [comment] }] });
		expect([refReads, textReads, authorReads]).toEqual([1, 1, 1]);
	});
});

describe("writeXlsx — bridge carry", () => {
	it("carries comments from a read workbook (openpyxl fixture) across the bridge", async () => {
		const { loadFixture } = await import("@openjsxl/fixtures");
		const before = await openXlsx(await loadFixture("openpyxl-comments.xlsx"));
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		expect(after.sheet("Notes").comments).toEqual([
			{ ref: "B2", author: "Ada", text: "check this figure" },
			{ ref: "C3", author: "Grace", text: "EMEA only" },
			{ ref: "D4", text: "no attribution" },
		]);
	});
});
