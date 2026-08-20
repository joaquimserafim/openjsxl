import { describe, expect, it } from "vitest";
import type { Comment } from "../../types";
import { commentsParts, vmlShape } from "../comments";

describe("commentsParts", () => {
	it("omits missing and empty comment collections", () => {
		expect(commentsParts("Data", undefined)).toBeUndefined();
		expect(commentsParts("Data", [])).toBeUndefined();
	});

	it("deduplicates logical authors before encoding Xstring text", () => {
		const parts = commentsParts("Data", [
			{ ref: "A1", author: "_x0041_", text: "_x0042_" },
			{ ref: "C3", author: "_x0041_", text: "  note  " },
			{ ref: "D4", text: "anonymous" },
		]);

		expect(parts?.commentsXml).toContain(
			"<authors><author>_x005F_x0041_</author><author></author></authors>",
		);
		expect(parts?.commentsXml).toContain(
			'<comment ref="A1" authorId="0" shapeId="0"><text><t>_x005F_x0042_</t></text></comment>',
		);
		expect(parts?.commentsXml).toContain(
			'<comment ref="C3" authorId="0" shapeId="0"><text><t xml:space="preserve">  note  </t></text></comment>',
		);
		expect(parts?.commentsXml).toContain('<comment ref="D4" authorId="1" shapeId="0">');
	});

	it("anchors each VML shape at the comment cell's zero-based row and column", () => {
		const parts = commentsParts("Data", [{ ref: "B2", author: "Ada", text: "Note" }]);

		expect(parts?.vmlXml).toContain('id="_x0000_s1025"');
		expect(parts?.vmlXml).toContain("<x:Row>1</x:Row><x:Column>1</x:Column>");
	});

	it.each([
		["non-array comments", {} as unknown as readonly Comment[], "comments must be an array"],
		[
			"unknown keys",
			[{ ref: "A1", text: "Note", typo: true }] as unknown as readonly Comment[],
			'comments[0] has an unknown property "typo"',
		],
		[
			"non-canonical references",
			[{ ref: "A0", text: "Note" }] as readonly Comment[],
			'comments[0].ref "A0" is not a canonical A1 cell',
		],
	])("rejects %s", (_label, comments, message) => {
		expect(() => commentsParts("Data", comments)).toThrow(message);
	});
});

describe("vmlShape", () => {
	it("uses the index for a unique shape id and z-index", () => {
		const xml = vmlShape(2, 4, 6);

		expect(xml).toContain('id="_x0000_s1027"');
		expect(xml).toContain("z-index:3");
		expect(xml).toContain("<x:Row>4</x:Row><x:Column>6</x:Column>");
	});
});
