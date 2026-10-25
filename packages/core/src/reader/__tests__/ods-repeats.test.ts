import { describe, expect, it } from "vitest";
import { MAX_GENERATED_TEXT, MAX_ODS_REPEATED_CELLS, parseOdsContent } from "../../ods/content";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { writeZip } from "../../writer/zip";
import { openOds } from "../ods";

// F11.57: repeats are expanded as runs, and only the cells a repeat creates beyond its element count
// against a document-wide budget, which fails typed when spent. Literal cells and literal text are
// never dropped or cut; only the extra spaces a text:s run expands to are capped.

const OPEN = "<office:document-content><office:body><office:spreadsheet>";
const CLOSE = "</office:spreadsheet></office:body></office:document-content>";
const FLOAT = '<table:table-cell office:value-type="float" office:value="1"';
const LITERAL_ROW = `<table:table-row>${`${FLOAT}/>`.repeat(100)}</table:table-row>`;

const budgetError = expect.objectContaining({
	name: "XlsxError",
	code: "part-too-large",
	message: expect.stringContaining("repeated cells exceed"),
});

describe("the repeat budget", () => {
	it("counts no literal cell, so a sheet larger than the budget reads every one", () => {
		const row = `<table:table-row>${`${FLOAT}/>`.repeat(300)}</table:table-row>`;
		const [sheet] = parseOdsContent(
			`${OPEN}<table:table table:name="S">${row}</table:table>${CLOSE}`,
			10,
		);
		expect(sheet?.cells.size).toBe(300);
		expect(sheet?.dimension).toBe("A1:KN1");
	});

	it("keeps literal cells that follow repeats at the edge of the budget", () => {
		// The most full-width rows of one value that fit: one more row would overrun the budget.
		const rows = Math.floor((MAX_ODS_REPEATED_CELLS + 1) / 16_384);
		expect(rows * 16_384 - 1).toBeLessThanOrEqual(MAX_ODS_REPEATED_CELLS);
		expect((rows + 1) * 16_384 - 1).toBeGreaterThan(MAX_ODS_REPEATED_CELLS);
		const repeated = `<table:table-row table:number-rows-repeated="${rows}">${FLOAT} table:number-columns-repeated="16384"/></table:table-row>`;
		const literal = `<table:table-row>${`${FLOAT}/>`.repeat(1500)}</table:table-row>`;
		const [sheet] = parseOdsContent(
			`${OPEN}<table:table table:name="S">${repeated}${literal}</table:table>${CLOSE}`,
		);
		expect(sheet?.cells.size).toBe(rows * 16_384 + 1500);
		expect(sheet?.cells.get(`BER${rows + 1}`)?.value).toBe(1);
	}, 60_000);

	it("fails typed on a small repeat bomb, before placing it", async () => {
		const content = `${OPEN}<table:table table:name="S"><table:table-row table:number-rows-repeated="1048576">${FLOAT} table:number-columns-repeated="16384"/></table:table-row></table:table>${CLOSE}`;
		const bytes = await writeZip([
			{
				name: "mimetype",
				data: new TextEncoder().encode("application/vnd.oasis.opendocument.spreadsheet"),
			},
			{ name: "content.xml", data: new TextEncoder().encode(content) },
		]);
		expect(bytes.length).toBeLessThan(1000);
		await expect(openOds(bytes)).rejects.toEqual(budgetError);
	});

	it("spans every sheet of the document", () => {
		// Each sheet alone creates 599 cells, within the 1,000 budget; the second one overruns it.
		const sheet = `<table:table table:name="S"><table:table-row table:number-rows-repeated="20">${FLOAT} table:number-columns-repeated="30"/></table:table-row></table:table>`;
		expect(parseOdsContent(`${OPEN}${sheet}${CLOSE}`, 1000)[0]?.cells.size).toBe(600);
		expect(() => parseOdsContent(`${OPEN}${sheet}${sheet}${CLOSE}`, 1000)).toThrow(budgetError);
	});

	it("places a row's literal cells once, however many stray row closes follow", () => {
		const [sheet] = parseOdsContent(
			`${OPEN}<table:table table:name="S">${LITERAL_ROW}${"</table:table-row>".repeat(50)}</table:table>${CLOSE}`,
			0,
		);
		expect(sheet?.cells.size).toBe(100);
	});

	it("places a row's literal cells once when rows are nested inside cells", () => {
		const depth = 50;
		const opens = "<table:table-row><table:table-cell>".repeat(depth);
		const closes = "</table:table-cell></table:table-row>".repeat(depth);
		const [sheet] = parseOdsContent(
			`${OPEN}<table:table table:name="S">${opens}${LITERAL_ROW}${closes}</table:table>${CLOSE}`,
			0,
		);
		expect(sheet?.cells.size).toBe(100);
	});

	it("still drops an empty repeat to the grid edge for free", () => {
		const tail =
			'<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>';
		const [sheet] = parseOdsContent(
			`${OPEN}<table:table table:name="S"><table:table-row>${FLOAT}/></table:table-row>${tail}</table:table>${CLOSE}`,
			0,
		);
		expect([sheet?.cells.size, sheet?.dimension]).toEqual([1, "A1:A1"]);
	});

	it("expands a repeated row in time linear in the rows it creates, up to 1M rows", () => {
		const ratio = bestDoublingRatio((n) => {
			const xml = `${OPEN}<table:table table:name="S"><table:table-row table:number-rows-repeated="${n}">${FLOAT}/></table:table-row></table:table>${CLOSE}`;
			return () => parseOdsContent(xml);
		}, 500_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);
});

describe("cell text", () => {
	const CELL_OPEN = `${OPEN}<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>`;
	const CELL_CLOSE = `</text:p></table:table-cell></table:table-row></table:table>${CLOSE}`;

	it("keeps literal text in full", () => {
		const [sheet] = parseOdsContent(`${CELL_OPEN}${"x".repeat(40_000)}${CELL_CLOSE}`);
		expect(String(sheet?.cells.get("A1")?.value).length).toBe(40_000);
	});

	it("keeps single spaces, tabs and line breaks in full", () => {
		const [sheet] = parseOdsContent(
			`${CELL_OPEN}${"a<text:tab/>".repeat(20_000)}${"b<text:s/>".repeat(20_000)}${"c<text:line-break/>".repeat(20_000)}${CELL_CLOSE}`,
		);
		expect(String(sheet?.cells.get("A1")?.value).length).toBe(120_000);
	});

	it("caps the extra spaces text:s runs expand to, keeping everything else", () => {
		const bomb = `${'<text:s text:c="1000000000"/>'.repeat(100)}${"<text:tab/>".repeat(10)}`;
		const [sheet] = parseOdsContent(`${CELL_OPEN}a${bomb}b${CELL_CLOSE}`);
		const value = String(sheet?.cells.get("A1")?.value);
		// Each run's first space is free; the extra spaces stop at the cap.
		expect(value.length).toBe(1 + 100 + MAX_GENERATED_TEXT + 10 + 1);
		expect([value.startsWith("a "), value.endsWith("\t\tb")]).toEqual([true, true]);
	});
});
