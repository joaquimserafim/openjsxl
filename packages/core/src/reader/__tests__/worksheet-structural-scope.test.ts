import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import type { DecodeContext } from "../../ooxml";
import type { Row } from "../../types";
import { writeZip } from "../../writer/zip";
import { openZip } from "../../zip";
import { openXlsx, streamSheetRows } from "../workbook";
import { parseCellStyles, readRows, streamRows } from "../worksheet";

// F11.22: style metadata counts only in worksheet/cols and worksheet/sheetData/row. Extension
// elements can reuse OOXML local names, so these examples deliberately place convincing lookalikes
// around the real containers and assert that values and styles still describe the same cells.

const dateStyleContext: DecodeContext = {
	sharedStrings: [],
	styles: {
		isDateStyle: (index) => index === 2,
		formatCode: (index) => (index === 2 ? "mm-dd-yy" : "General"),
		cellStyle: () => undefined,
	},
};

/**
 * Rebuild the fixture after changing only sheet1 XML and verify the mutation actually landed.
 * Examples: a mutation with all required sentinels present and forbidden sentinels absent returns
 * new XLSX bytes; unchanged XML, a missing required marker, a surviving forbidden marker, or a
 * fixture without sheet1 fails its assertion instead of returning a misleading test workbook.
 */
async function rewriteWorksheet(
	mutate: (xml: string) => string,
	presentSentinels: readonly string[],
	absentSentinels: readonly string[] = [],
): Promise<Uint8Array> {
	const source = openZip(await loadFixture("col-row-styles.xlsx"));
	const decoder = new TextDecoder();
	const encoder = new TextEncoder();
	const parts: { name: string; data: Uint8Array }[] = [];
	let rewritten = false;
	for (const name of source.entries.keys()) {
		let data = await source.read(name);
		if (name === "xl/worksheets/sheet1.xml") {
			rewritten = true;
			const original = decoder.decode(data);
			const changed = mutate(original);
			expect(changed).not.toBe(original); // fail if fixture spelling stops injecting the attack
			for (const sentinel of presentSentinels) expect(changed).toContain(sentinel);
			for (const sentinel of absentSentinels) expect(changed).not.toContain(sentinel);
			data = encoder.encode(changed);
		}
		parts.push({ name, data });
	}
	expect(rewritten).toBe(true);
	return writeZip(parts);
}

/**
 * Materialize streamed rows without changing their order.
 * Examples: an async source yielding rows 1 then 3 returns `[row1, row3]`; an empty source returns
 * `[]`; a source error is propagated to the test.
 */
async function collectRows(rows: AsyncIterable<Row>): Promise<Row[]> {
	const out: Row[] = [];
	for await (const row of rows) out.push(row);
	return out;
}

/**
 * Encode XML and expose it as fixed-size parser chunks; callers supply a positive `size`.
 * Examples: input `("abcdef", 2)` yields bytes for `"ab"`, `"cd"`, `"ef"`; `("abc", 2)` yields
 * `"ab"`, then the shorter `"c"`; empty XML yields no chunks.
 */
async function* xmlChunks(xml: string, size: number): AsyncGenerator<Uint8Array> {
	const bytes = new TextEncoder().encode(xml);
	for (let offset = 0; offset < bytes.length; offset += size) {
		yield bytes.subarray(offset, offset + size);
	}
}

describe("worksheet style structural scope — parser matrix", () => {
	it("ignores fake columns, rows, and cells before, after, and under extensions", () => {
		const xml =
			'<worksheet><col min="1" max="1" style="7"/>' +
			'<row r="8"><c r="A8" s="7"/></row><c r="A9" s="7"/>' +
			'<extLst><ext><cols><col min="1" max="1" style="7"/></cols>' +
			'<sheetData><row r="7"><c r="A7" s="7"/></row></sheetData></ext></extLst>' +
			'<cols><ext><col min="1" max="1" style="7"/></ext>' +
			'<col min="1" max="1" style="2"/></cols><col min="1" max="1" style="7"/>' +
			'<sheetData><row r="1"><ext><row r="5"><c r="A5" s="7"/></row></ext>' +
			'<c r="A1"><v>45000</v><extLst><ext><c r="A1" s="7"/></ext></extLst></c></row>' +
			'<ext><row r="1"><c r="A1" s="7"/></row></ext></sheetData>' +
			'<col min="1" max="1" style="7"/><row r="6"><c r="A6" s="7"/></row>' +
			'<extLst><ext><row r="1"><c r="A1" s="7"/></row></ext></extLst></worksheet>';

		// Input has many style-7 lookalikes; output contains only canonical A1 with style 2.
		expect(parseCellStyles(xml)).toEqual(new Map([["A1", 2]]));
		expect([...readRows(xml, dateStyleContext)]).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "date", value: expect.any(Date) }] },
		]);
	});

	it("clears row scope for self-closing and closed rows", () => {
		const xml =
			'<worksheet><cols/><col min="1" max="1" style="7"/>' +
			'<sheetData/><row r="9"><c r="A9" s="7"/></row>' +
			'<sheetData><row r="1" s="7" customFormat="1"/><c r="A1" s="7"/>' +
			'<row r="2"><c r="A2" s="2"/></row><c r="B2" s="7"/></sheetData>' +
			'<sheetData/><c r="A3" s="7"/></worksheet>';

		// A1/B2 are outside an active row; the self-closing row itself remains a real empty row.
		expect(parseCellStyles(xml)).toEqual(new Map([["A2", 2]]));
		expect([...readRows(xml, dateStyleContext)]).toEqual([
			{ index: 1, cells: [] },
			{ index: 2, cells: [{ ref: "A2", type: "empty", value: null }] },
		]);
	});

	it("keeps prefixes and truncated canonical row/cell recovery", async () => {
		const xml =
			'<x:worksheet><x:cols><x:col min="1" max="1" style="2"/></x:cols>' +
			'<x:sheetData><x:row r="1"><x:c r="A1"><x:v>45000';

		// EOF is the output boundary: the readable open cell and row are flushed once.
		expect(parseCellStyles(xml)).toEqual(new Map([["A1", 2]]));
		const rows = [...readRows(xml, dateStyleContext)];
		expect(rows).toHaveLength(1);
		expect(rows[0]?.cells[0]).toMatchObject({ ref: "A1", type: "date" });
		expect(await collectRows(streamRows(xmlChunks(xml, 3), dateStyleContext))).toEqual(rows);
	});

	it("closes canonical scope at a misnested worksheet root", () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1" s="2"/></row></worksheet>' +
			'<row r="2"><c r="A2" s="7"/></row></sheetData>' +
			'<sheetData><row r="3"><c r="A3" s="7"/></row></sheetData>';

		expect(parseCellStyles(xml)).toEqual(new Map([["A1", 2]]));
		expect([...readRows(xml, dateStyleContext)]).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "empty", value: null }] },
		]);
	});

	it("recovers canonical cell and row closes across unclosed extension markup", () => {
		const cellXml =
			'<worksheet><sheetData><row r="1"><c r="A1" s="2"><v>1</v><ext></c>' +
			'<c r="B1" s="2"><v>2</v></c></row></sheetData></worksheet>';
		expect(parseCellStyles(cellXml)).toEqual(
			new Map([
				["A1", 2],
				["B1", 2],
			]),
		);
		expect([...readRows(cellXml, { sharedStrings: [] })]).toEqual([
			{
				index: 1,
				cells: [
					{ ref: "A1", type: "number", value: 1 },
					{ ref: "B1", type: "number", value: 2 },
				],
			},
		]);

		const rowXml =
			'<worksheet><sheetData><row r="1"><c r="A1" s="2"/><ext></row>' +
			'<row r="2"><c r="A2" s="2"/></row></sheetData></worksheet>';
		expect(parseCellStyles(rowXml)).toEqual(
			new Map([
				["A1", 2],
				["A2", 2],
			]),
		);
		expect([...readRows(rowXml, { sharedStrings: [] })]).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "empty", value: null }] },
			{ index: 2, cells: [{ ref: "A2", type: "empty", value: null }] },
		]);
	});
});

describe("worksheet style structural scope — public accessors", () => {
	it.each([
		{
			name: "exact extension A1 cannot replace the real General-style cell",
			ref: "A1",
			type: "number",
			format: "General",
			styleFormat: undefined,
			expectedValue: 1,
			absentRefs: [] as string[],
			presentSentinels: ['<c r="A1" s="2"/>'],
			absentSentinels: [] as string[],
			mutate: (xml: string) =>
				xml.replace(
					"</worksheet>",
					'<extLst><ext><row r="1"><c r="A1" s="2"/></row></ext></extLst></worksheet>',
				),
		},
		{
			name: "extension cell cannot overwrite real B1",
			ref: "B1",
			type: "date",
			format: "mm-dd-yy",
			styleFormat: "mm-dd-yy",
			expectedValue: "2020-01-01T00:00:00.000Z",
			absentRefs: [] as string[],
			presentSentinels: ['<c r="B1" s="2"/>'],
			absentSentinels: [] as string[],
			mutate: (xml: string) =>
				xml.replace(
					"</worksheet>",
					'<extLst><ext><row r="1"><c r="B1" s="2"/></row></ext></extLst></worksheet>',
				),
		},
		{
			name: "extension column cannot outrank canonical column B",
			ref: "B1",
			type: "date",
			format: "mm-dd-yy",
			styleFormat: "mm-dd-yy",
			expectedValue: "2020-01-01T00:00:00.000Z",
			absentRefs: [] as string[],
			presentSentinels: ['<extLst><ext><col min="2" max="2" style="2"/>'],
			absentSentinels: [] as string[],
			mutate: (xml: string) =>
				xml.replace(
					"<cols>",
					'<extLst><ext><col min="2" max="2" style="2"/></ext></extLst><cols>',
				),
		},
		{
			name: "combined before, inside, and after lookalikes stay absent",
			ref: "B1",
			type: "date",
			format: "mm-dd-yy",
			styleFormat: "mm-dd-yy",
			expectedValue: "2020-01-01T00:00:00.000Z",
			absentRefs: ["A88", "A89", "A90", "A91"],
			presentSentinels: ["A88", "A89", "A90", "A91", "<sheetData/>", "<cols><ext><col"],
			absentSentinels: [] as string[],
			mutate: (xml: string) =>
				xml
					.replace("<cols>", '<cols><ext><col min="2" max="2" style="2"/></ext>')
					.replace(
						"</cols><sheetData>",
						'</cols><col min="2" max="2" style="2"/><sheetData/>' +
							'<row r="88"><c r="A88" s="2"/></row><sheetData>' +
							'<ext><row r="89"><c r="A89" s="2"/></row></ext>',
					)
					.replace(
						"</sheetData></worksheet>",
						'</sheetData><row r="90"><c r="A90" s="2"/></row>' +
							'<c r="A91" s="2"/><sheetData/></worksheet>',
					),
		},
		{
			name: "prefixed canonical style elements retain the same model",
			ref: "B1",
			type: "date",
			format: "mm-dd-yy",
			styleFormat: "mm-dd-yy",
			expectedValue: "2020-01-01T00:00:00.000Z",
			absentRefs: [] as string[],
			presentSentinels: [
				"<x:worksheet",
				"<x:cols",
				"<x:col ",
				"<x:sheetData",
				"<x:row",
				"<x:c ",
				"<x:v>",
			],
			absentSentinels: [] as string[],
			mutate: (xml: string) =>
				xml
					.replace(
						"<worksheet ",
						'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
					)
					.replace("</worksheet>", "</x:worksheet>")
					.replaceAll("<cols", "<x:cols")
					.replaceAll("</cols>", "</x:cols>")
					.replaceAll("<col ", "<x:col ")
					.replaceAll("<sheetData", "<x:sheetData")
					.replaceAll("</sheetData>", "</x:sheetData>")
					.replaceAll("<row", "<x:row")
					.replaceAll("</row>", "</x:row>")
					.replaceAll("<c ", "<x:c ")
					.replaceAll("</c>", "</x:c>")
					.replaceAll("<v>", "<x:v>")
					.replaceAll("</v>", "</x:v>"),
		},
		{
			name: "truncated final cell and row remain readable",
			ref: "B3",
			type: "date",
			format: "mm-dd-yy",
			styleFormat: "mm-dd-yy",
			expectedValue: "2020-09-26T00:00:00.000Z",
			absentRefs: [] as string[],
			presentSentinels: ['<c r="B3" s="1"><v>44100</v>'],
			absentSentinels: ["</worksheet>"],
			mutate: (xml: string) =>
				xml.replace(
					'<c r="B3" s="1"><v>44100</v></c></row></sheetData></worksheet>',
					'<c r="B3" s="1"><v>44100</v>',
				),
		},
	])(
		"keeps cell, rows, style, and numberFormat aligned: $name",
		async ({
			mutate,
			ref,
			type,
			format,
			styleFormat,
			expectedValue,
			absentRefs,
			presentSentinels,
			absentSentinels,
		}) => {
			const bytes = await rewriteWorksheet(mutate, presentSentinels, absentSentinels);
			const sheet = (await openXlsx(bytes)).sheet("Sheet1");

			const direct = sheet.cell(ref);
			expect(direct.type).toBe(type);
			if (typeof expectedValue === "number") expect(direct.value).toBe(expectedValue);
			else {
				expect(direct.value).toBeInstanceOf(Date);
				if (direct.value instanceof Date)
					expect(direct.value.toISOString()).toBe(expectedValue);
			}
			const style = sheet.style(ref);
			if (styleFormat === undefined) expect(style).toBeUndefined();
			else expect(style?.numberFormat).toBe(styleFormat);
			expect(sheet.numberFormat(ref)).toBe(format);

			const buffered = await collectRows(sheet.rows());
			const streamed = await collectRows(streamSheetRows(bytes));
			expect(buffered).toEqual(streamed);
			const cells = buffered.flatMap((row) => row.cells);
			expect(cells.find((cell) => cell.ref === ref)).toEqual(direct);
			expect(cells.filter((cell) => absentRefs.includes(cell.ref))).toEqual([]);
			for (const absentRef of absentRefs) {
				expect(sheet.cell(absentRef)).toEqual({
					ref: absentRef,
					type: "empty",
					value: null,
				});
				expect(sheet.style(absentRef)).toBeUndefined();
				expect(sheet.numberFormat(absentRef)).toBe("General");
			}
		},
	);
});
