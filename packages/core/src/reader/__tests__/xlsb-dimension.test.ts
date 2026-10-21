import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { streamXlsx, workbookToInput, writeXlsx } from "../../writer";
import { openXlsx } from "../workbook";
import { openXlsb } from "../xlsb";

// F11.16: controlled, synthetic BIFF12 records, read from the committed archive verbatim.
// The provenance in fixtures/data/README.md explicitly distinguishes this from Excel output.
const fixture = async () => openXlsb(await loadFixture("xlsb-dimensions.xlsb"));

describe("XLSB dimensions — committed fixture", () => {
	it("does not invent A1 for an empty zero-sentinel sheet", async () => {
		const sheet = (await fixture()).sheet("Empty");
		expect(sheet.dimension).toBeUndefined();
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "empty", value: null });
		const rows = [];
		for await (const row of sheet.rows()) rows.push(row);
		expect(rows).toEqual([]);
	});

	it("retains A1 when the same zero coordinates accompany an A1 value", async () => {
		const sheet = (await fixture()).sheet("A1Value");
		expect(sheet.dimension).toBe("A1");
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "number", value: 42 });
	});

	it("retains A1 for a styled blank without fabricating a cell value", async () => {
		const sheet = (await fixture()).sheet("A1Blank");
		expect(sheet.dimension).toBe("A1");
		expect(sheet.numberFormat("A1")).toBe("mm-dd-yy");
		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "empty", value: null });
	});

	it("preserves an ordinary explicit range", async () => {
		const sheet = (await fixture()).sheet("Ordinary");
		expect(sheet.dimension).toBe("B4:B7");
		expect(sheet.cell("B4").value).toBe(10);
		expect(sheet.cell("B7").value).toBe(20);
	});

	it("treats an all-max invalid range as absent", async () => {
		const sheet = (await fixture()).sheet("InvalidMax");
		expect(sheet.dimension).toBeUndefined();
		expect(sheet.cell("XFD1048576").type).toBe("empty");
	});

	it("reconciles zero coordinates with sparse values and a styled blank away from A1", async () => {
		const sheet = (await fixture()).sheet("ZeroBounds");
		expect(sheet.dimension).toBe("B4:D7");
		expect(sheet.cell("B4").value).toBe(10);
		expect(sheet.cell("B7").value).toBe(20);
		expect(sheet.cell("D7").type).toBe("empty");
		expect(sheet.numberFormat("D7")).toBe("mm-dd-yy");
	});

	it("retains dimensions in the pre-existing synthetic XLSB corpus", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-basic.xlsb"));
		expect(wb.sheet("Sheet1").dimension).toBe("A1:G2");
		expect(wb.sheet("Sheet2").dimension).toBe("A1");
		expect(wb.sheet("Hidden").dimension).toBe("A1");
	});

	it.each([
		{ name: "RowDefault", dimension: "A1" },
		{ name: "ColumnDefault", dimension: "A1" },
		{ name: "MixedDefaults", dimension: "D5" },
	])("counts default formatting in $name without inventing a cell", async ({
		name,
		dimension,
	}) => {
		const sheet = (await fixture()).sheet(name);
		expect(sheet.dimension).toBe(dimension);
		expect(sheet.cell(dimension).type).toBe("empty");
		const rows = [];
		for await (const row of sheet.rows()) rows.push(row);
		expect(rows).toEqual([]);
	});

	it.each([
		"buffered",
		"streaming",
	])("preserves empty sheets and values through the %s bridge", async (mode) => {
		const input = await workbookToInput(await fixture());
		const bytes =
			mode === "buffered"
				? await writeXlsx(input)
				: new Uint8Array(await new Response(streamXlsx(input)).arrayBuffer());
		const rewritten = await openXlsx(bytes);
		for (const name of ["Empty", "InvalidMax"]) {
			const sheet = rewritten.sheet(name);
			// The buffered XLSX writer's existing empty-sheet convention is dimension A1;
			// streaming omits dimensions. Neither convention creates a cell or row.
			expect(sheet.dimension).toBe(mode === "buffered" ? "A1" : undefined);
			expect(sheet.cell("A1").type).toBe("empty");
			const rows = [];
			for await (const row of sheet.rows()) rows.push(row);
			expect(rows).toEqual([]);
		}
		expect(rewritten.sheet("A1Value").cell("A1").value).toBe(42);
		for (const name of ["Ordinary", "ZeroBounds"]) {
			expect(rewritten.sheet(name).cell("B4").value).toBe(10);
			expect(rewritten.sheet(name).cell("B7").value).toBe(20);
		}
		// Streaming omits dimensions; XLSB styled blanks are not surfaced in rows(), so their
		// styles/bounds are not carried by the existing bridge. Only populated buffered bounds apply.
		if (mode === "buffered") {
			expect(rewritten.sheet("A1Value").dimension).toBe("A1");
			expect(rewritten.sheet("Ordinary").dimension).toBe("B4:B7");
		}
	});
});
