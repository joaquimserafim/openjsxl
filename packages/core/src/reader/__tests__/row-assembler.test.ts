import { describe, expect, it } from "vitest";
import type { DecodeContext, StyleTable } from "../../ooxml";
import type { Row } from "../../types";
import type { XmlToken } from "../../xml";
import { createRowAssembler } from "../row-assembler";

// These token helpers bypass the XML tokenizer so each test exercises only row-assembly state.
const open = (
	name: string,
	attrs: Readonly<Record<string, string>> = {},
	selfClosing = false,
): XmlToken => ({ kind: "open", name, attrs, selfClosing });
const text = (value: string): XmlToken => ({ kind: "text", value });
const close = (name: string): XmlToken => ({ kind: "close", name });

// Feed a complete example and include the assembler's intentional end-of-input recovery.
function assemble(tokens: readonly XmlToken[], ctx: DecodeContext = { sharedStrings: [] }): Row[] {
	const assembler = createRowAssembler(ctx);
	const rows: Row[] = [];
	assembler.push(open("worksheet"));
	for (const token of tokens) rows.push(...assembler.push(token));
	rows.push(...assembler.flush());
	return rows;
}

function valueCell(attrs: Readonly<Record<string, string>>, value: string): XmlToken[] {
	return [open("c", attrs), open("v"), text(value), close("v"), close("c")];
}

describe("createRowAssembler — row and cell state", () => {
	it("assembles only rows inside sheetData and accepts namespace prefixes", () => {
		const assembler = createRowAssembler({ sharedStrings: [] });
		expect(assembler.push(open("x:worksheet"))).toEqual([]);
		expect(assembler.push(open("x:row", { r: "9" }, true))).toEqual([]);
		expect(assembler.push(open("x:sheetData", {}, true))).toEqual([]);
		expect(assembler.push(open("x:row", { r: "8" }, true))).toEqual([]);
		expect(assembler.push(open("x:sheetData"))).toEqual([]);
		expect(assembler.push(open("x:row", { r: "2" }))).toEqual([]);
		expect(assembler.push(open("x:c", { r: "B2" }, true))).toEqual([]);

		// A normal row is complete only when its closing token arrives.
		expect(assembler.push(close("x:row"))).toEqual([
			{
				index: 2,
				cells: [{ ref: "B2", type: "empty", value: null }],
			},
		]);
		expect(assembler.flush()).toEqual([]);
	});

	it("numbers rows positionally when r is missing or invalid", () => {
		const rows = assemble([
			open("sheetData"),
			open("row", { r: "7" }, true),
			open("row", { r: "0" }, true),
			open("row", { r: "bad" }, true),
			open("row", {}, true),
		]);
		expect(rows.map((row) => row.index)).toEqual([7, 8, 9, 10]);
	});

	it("can emit an unclosed row and a new self-closing row from one token", () => {
		const assembler = createRowAssembler({ sharedStrings: [] });
		assembler.push(open("worksheet"));
		assembler.push(open("sheetData"));
		assembler.push(open("row", { r: "1" }));
		assembler.push(open("c", { r: "A1" }, true));

		// The new row first recovers row 1, then immediately completes its own `<row/>`.
		expect(assembler.push(open("row", { r: "2" }, true))).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "empty", value: null }] },
			{ index: 2, cells: [] },
		]);
	});

	it("keeps explicit refs and continues missing refs from the last valid column", () => {
		const rows = assemble([
			open("sheetData"),
			open("row", { r: "5" }),
			open("c", { r: "C99" }, true),
			open("c", {}, true),
			open("c", { r: "bad" }, true),
			open("c", {}, true),
			close("row"),
			close("sheetData"),
		]);

		// Explicit refs stay verbatim; positional refs use the containing row (5).
		expect(rows[0]?.cells.map((cell) => cell.ref)).toEqual(["C99", "D5", "bad", "E5"]);
	});

	it("distinguishes missing values from present empty values and isolates channels", () => {
		const rows = assemble([
			open("sheetData"),
			open("row", { r: "1" }),
			open("c", { r: "A1" }, true),
			open("c", { r: "B1", t: "str" }),
			open("v", {}, true),
			close("c"),
			open("c", { r: "C1", t: "inlineStr" }),
			open("is", {}, true),
			close("c"),
			open("c", { r: "D1", t: "b" }),
			open("is"),
			open("t"),
			text("ignored"),
			close("t"),
			close("is"),
			open("v"),
			text("1"),
			close("v"),
			close("c"),
			open("c", { r: "E1", t: "inlineStr" }),
			open("is"),
			open("t"),
			text("kept"),
			close("t"),
			close("is"),
			open("v"),
			text("99"),
			close("v"),
			close("c"),
			close("row"),
		]);

		expect(rows[0]?.cells).toEqual([
			{ ref: "A1", type: "empty", value: null },
			{ ref: "B1", type: "string", value: "" },
			{ ref: "C1", type: "string", value: "" },
			{ ref: "D1", type: "boolean", value: true },
			{ ref: "E1", type: "string", value: "kept" },
		]);
	});

	it("recovers unclosed cells and rows, then flushes EOF only once", () => {
		const assembler = createRowAssembler({ sharedStrings: [] });
		assembler.push(open("worksheet"));
		assembler.push(open("sheetData"));
		assembler.push(open("row", { r: "1" }));
		for (const token of valueCell({ r: "A1" }, "1").slice(0, -1)) assembler.push(token);

		// Opening B1 finishes the missing `</c>` for A1.
		assembler.push(open("c", { r: "B1" }, true));
		expect(assembler.push(open("row", { r: "2" }))).toEqual([
			{
				index: 1,
				cells: [
					{ ref: "A1", type: "number", value: 1 },
					{ ref: "B1", type: "empty", value: null },
				],
			},
		]);
		assembler.push(open("c", { r: "A2" }, true));

		// EOF preserves the otherwise readable final row; another flush has no duplicate.
		expect(assembler.flush()).toEqual([
			{ index: 2, cells: [{ ref: "A2", type: "empty", value: null }] },
		]);
		expect(assembler.flush()).toEqual([]);
	});

	it("finishes buffered inline text when sheetData closes early", () => {
		const assembler = createRowAssembler({ sharedStrings: [] });
		assembler.push(open("worksheet"));
		assembler.push(open("sheetData"));
		assembler.push(open("row", { r: "3" }));
		assembler.push(open("c", { r: "A3", t: "inlineStr" }));
		assembler.push(open("is"));
		assembler.push(open("t"));
		assembler.push(text("tail"));

		// Missing `</t></is></c></row>` still keeps the readable value at sheet-data close.
		expect(assembler.push(close("sheetData"))).toEqual([
			{ index: 3, cells: [{ ref: "A3", type: "string", value: "tail" }] },
		]);
		expect(assembler.flush()).toEqual([]);
	});
});

describe("createRowAssembler — text and styles", () => {
	it("decodes split Xstring text once per run and excludes phonetic guides", () => {
		const rows = assemble([
			open("sheetData"),
			open("row", { r: "1" }),
			open("c", { r: "A1", t: "inlineStr" }),
			open("is"),
			open("t"),
			text("_x00"),
			text("41_"),
			close("t"),
			close("is"),
			close("c"),
			open("c", { r: "B1", t: "inlineStr" }),
			open("is"),
			open("t"),
			text("_x00"),
			close("t"),
			open("t"),
			text("41_"),
			close("t"),
			close("is"),
			close("c"),
			open("c", { r: "C1", t: "inlineStr" }),
			open("is"),
			open("t"),
			text("漢字"),
			close("t"),
			open("rPh"),
			open("t"),
			text("かんじ"),
			close("t"),
			close("rPh"),
			close("is"),
			close("c"),
			...valueCell({ r: "D1", t: "str" }, "a_x000A_b"),
			...valueCell({ r: "E1", t: "e" }, "_x0041_"),
			close("row"),
		]);

		expect(rows[0]?.cells.map((cell) => cell.value)).toEqual([
			"A",
			"_x0041_",
			"漢字",
			"a\nb",
			"_x0041_",
		]);
	});

	it("uses cell, row, and column styles in that order", () => {
		const seen: Array<number | undefined> = [];
		const styles: StyleTable = {
			isDateStyle: (index) => {
				seen.push(index);
				return false;
			},
			formatCode: () => undefined,
			cellStyle: () => undefined,
		};
		assemble(
			[
				open("cols"),
				open("col", { min: "1", max: "4", style: "7" }, true),
				close("cols"),
				open("sheetData"),
				open("row", { r: "1", s: "5", customFormat: "1" }),
				...valueCell({ r: "A1" }, "1"),
				...valueCell({ r: "B1", s: "3" }, "2"),
				...valueCell({ r: "C1", s: "bad" }, "3"),
				close("row"),
				open("row", { r: "2" }),
				...valueCell({ r: "A2" }, "4"),
				close("row"),
			],
			{ sharedStrings: [], styles },
		);

		// An invalid own style still blocks inheritance, so C1 deliberately resolves to style 0.
		expect(seen).toEqual([5, 3, undefined, 7]);
	});
});
