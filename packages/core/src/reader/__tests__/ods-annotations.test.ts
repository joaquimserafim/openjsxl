import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { parseOdsContent } from "../../ods";
import { workbookToInput, writeXlsx } from "../../writer";
import { openOds } from "../ods";
import { openXlsx } from "../workbook";

// F11.56: a cell's comment (`office:annotation`) and the shapes anchored in it (`draw:*`) are not cell
// text. Their paragraphs, links and text never reach the cell's value or hyperlink.

const fixture = await loadFixture("ods-annotations.ods");
const REFS = ["A1", "B1", "C1", "D1", "E1", "F1", "A2", "B2", "C2", "D2", "E2"];
const EXPECTED = [
	{ ref: "A1", type: "string", value: "Revenue" },
	{ ref: "B1", type: "empty", value: null },
	{ ref: "C1", type: "error", value: "#DIV/0!" },
	{ ref: "D1", type: "string", value: "plain" },
	{ ref: "E1", type: "string", value: "anchored" },
	{ ref: "F1", type: "string", value: "shape" },
	{ ref: "A2", type: "string", value: "line 1\nline 2" },
	{ ref: "B2", type: "string", value: "site" },
	{ ref: "C2", type: "number", value: 5 },
	{ ref: "D2", type: "string", value: "solo" },
	{ ref: "E2", type: "string", value: "someday" },
];

describe("openOds — comments and anchored shapes (ods-annotations.ods)", () => {
	it("reads each cell as its own text, without its comment or shapes", async () => {
		const sheet = (await openOds(fixture)).sheet("Notes");
		expect(REFS.map((ref) => sheet.cell(ref))).toEqual(EXPECTED);
	});

	it("takes a cell's hyperlink from the cell only, never from its comment", async () => {
		const sheet = (await openOds(fixture)).sheet("Notes");
		expect(sheet.hyperlinks).toEqual([{ ref: "B2", target: "https://cell.example/" }]);
	});

	it("bridges to xlsx with the same cells and links", async () => {
		const bridged = (
			await openXlsx(await writeXlsx(await workbookToInput(await openOds(fixture))))
		).sheet("Notes");
		// The writer has no error-cell form: an error without a formula bridges as its text.
		expect(REFS.map((ref) => bridged.cell(ref))).toEqual(
			EXPECTED.map((cell) => (cell.type === "error" ? { ...cell, type: "string" } : cell)),
		);
		expect(bridged.hyperlinks).toEqual([{ ref: "B2", target: "https://cell.example/" }]);
	});
});

describe("parseOdsContent — ignored containers", () => {
	const OPEN =
		'<office:document-content><office:body><office:spreadsheet><table:table table:name="S"><table:table-row>';
	const CLOSE =
		"</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>";

	it("skips a comment inside the cell's paragraph and keeps counting the cell's own paragraphs", () => {
		const [sheet] = parseOdsContent(
			OPEN +
				'<table:table-cell office:value-type="string"><text:p>Rev<office:annotation><text:p>x</text:p><text:p>y</text:p></office:annotation>enue</text:p><text:p>more</text:p></table:table-cell>' +
				'<table:table-cell office:value-type="string"><text:p>next</text:p></table:table-cell>' +
				CLOSE,
		);
		expect([sheet?.cells.get("A1")?.value, sheet?.cells.get("B1")?.value]).toEqual([
			"Revenue\nmore",
			"next",
		]);
	});

	it("skips shapes nested inside a comment", () => {
		const [sheet] = parseOdsContent(
			OPEN +
				'<table:table-cell office:value-type="string"><office:annotation><draw:frame><draw:text-box><text:p>deep</text:p></draw:text-box></draw:frame><text:p>c</text:p></office:annotation><text:p>v</text:p></table:table-cell>' +
				CLOSE,
		);
		expect(sheet?.cells.get("A1")?.value).toBe("v");
	});

	it("does not carry an unclosed comment into the next cell", () => {
		const [sheet] = parseOdsContent(
			OPEN +
				'<table:table-cell office:value-type="string"><office:annotation><text:p>x</text:p></table:table-cell>' +
				'<table:table-cell office:value-type="string"><text:p>ok</text:p></table:table-cell>' +
				CLOSE,
		);
		expect(sheet?.cells.get("B1")?.value).toBe("ok");
	});

	it("skips a shape written before the cell's paragraph too", () => {
		const [sheet] = parseOdsContent(
			OPEN +
				'<table:table-cell office:value-type="string"><draw:frame><draw:text-box><text:p>frame</text:p></draw:text-box></draw:frame><text:p>cell</text:p></table:table-cell>' +
				CLOSE,
		);
		expect(sheet?.cells.get("A1")?.value).toBe("cell");
	});

	it("never makes a cell from cell markup nested inside a comment", () => {
		const [sheet] = parseOdsContent(
			OPEN +
				'<table:table-cell office:value-type="string"><office:annotation><table:table-cell office:value-type="string"><text:p>INJ</text:p></table:table-cell></office:annotation><text:p>v</text:p></table:table-cell>' +
				'<table:table-cell office:value-type="string"><text:p>w</text:p></table:table-cell>' +
				CLOSE,
		);
		expect([...(sheet?.cells.values() ?? [])].map((c) => [c.ref, c.value])).toEqual([
			["A1", "v"],
			["B1", "w"],
		]);
	});
});
