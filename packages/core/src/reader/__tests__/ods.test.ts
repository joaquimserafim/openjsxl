import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { parseOdsContent } from "../../ods";
import { workbookToInput, writeXlsx } from "../../writer";
import { openOds } from "../ods";
import { openXlsx } from "../workbook";

// A minimal content.xml wrapper for unit-testing the parser directly (namespaces are matched by
// literal prefix, so the xmlns decls are unnecessary here).
const odsDoc = (body: string) =>
	`<office:document-content><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`;

// F7.1 — the .ods reader. Ground-truth values are cross-checked against python-calamine (an
// independent ODF reader; openpyxl can't read .ods): see the fixture provenance in data/README.md.
// odf-basic.ods is a real odfpy-authored file; ods-*.ods are crafted (generator) for edge/reject
// shapes LibreOffice would emit but odfpy won't.

describe("openOds — real-producer value matrix (odf-basic.ods)", () => {
	it("lists sheets in document order", async () => {
		const wb = await openOds(await loadFixture("odf-basic.ods"));
		expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet2"]);
	});

	it("reads every explicit value type", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		expect(s.cell("A1")).toEqual({ ref: "A1", type: "string", value: "hello" });
		expect(s.cell("B1")).toEqual({ ref: "B1", type: "number", value: 42 });
		expect(s.cell("C1")).toEqual({ ref: "C1", type: "number", value: -3.5 });
		expect(s.cell("D1")).toEqual({ ref: "D1", type: "boolean", value: true });
		expect(s.cell("E1")).toEqual({ ref: "E1", type: "boolean", value: false });
		// percentage + currency read as their raw numeric value (office:value is authoritative).
		expect(s.cell("H1")).toEqual({ ref: "H1", type: "number", value: 0.25 });
		expect(s.cell("I1")).toEqual({ ref: "I1", type: "number", value: 19.99 });
	});

	it("maps ODF date/date-time strings to the same UTC instant as the xlsx path", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		const f1 = s.cell("F1");
		const g1 = s.cell("G1");
		expect(f1.type).toBe("date");
		expect(g1.type).toBe("date");
		expect((f1.value as Date).getTime()).toBe(Date.UTC(2024, 0, 15));
		expect((g1.value as Date).getTime()).toBe(Date.UTC(2024, 0, 15, 13, 30, 0));
	});

	it("reads a formula cell as its cached value (formula text is not carried)", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		expect(s.cell("A2")).toEqual({ ref: "A2", type: "number", value: 84 });
		expect(s.formula("A2")).toBeUndefined();
	});

	it("captures a cell hyperlink from <text:a> and keeps the link text as the value", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		expect(s.cell("B2")).toEqual({ ref: "B2", type: "string", value: "apples" });
		expect(s.hyperlinks).toEqual([{ ref: "B2", target: "https://example.com/apples" }]);
	});

	it("reads a covered-cell merge as an A1 range", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		expect(s.mergedCells).toEqual(["A4:B5"]);
	});

	it("synthesizes a dimension from the populated cells", async () => {
		const wb = await openOds(await loadFixture("odf-basic.ods"));
		expect(wb.sheet("Sheet1").dimension).toBe("A1:I4");
		expect(wb.sheet("Sheet2").cell("A1").value).toBe("second");
		expect(wb.sheet("Sheet2").cell("B1").value).toBe(100);
	});

	it("streams rows in ascending row/column order", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet2");
		const rows = [];
		for await (const row of s.rows()) rows.push(row);
		expect(rows).toHaveLength(1);
		expect(rows[0]?.index).toBe(1);
		expect(rows[0]?.cells.map((c) => c.value)).toEqual(["second", 100]);
	});
});

describe("openOds — repeats, merges, visibility (ods-edge.ods)", () => {
	it("materializes a value repeated across columns", async () => {
		const s = (await openOds(await loadFixture("ods-edge.ods"))).sheet("Data");
		expect(s.cell("A2").value).toBe(7);
		expect(s.cell("B2").value).toBe(7);
		expect(s.cell("C2").value).toBe(7);
		expect(s.cell("D2").type).toBe("empty");
	});

	it("drops the empty repeat-to-grid-edge tail (bomb): dimension stays the used range", async () => {
		// The fixture pads to column 16384 and repeats an empty row ~1e6 times; an unbounded reader
		// would materialize ~2^34 cells. A correct reader returns instantly with the real used range.
		const s = (await openOds(await loadFixture("ods-edge.ods"))).sheet("Data");
		expect(s.dimension).toBe("A1:F3");
		expect(s.cell("A1").value).toBe("hello");
		expect(s.cell("F1")).toEqual({ ref: "F1", type: "string", value: "link" });
		expect(s.hyperlinks).toEqual([{ ref: "F1", target: "https://example.com/" }]);
	});

	it("reads a 2×2 covered-cell merge", async () => {
		const s = (await openOds(await loadFixture("ods-edge.ods"))).sheet("Data");
		expect(s.cell("A3").value).toBe("merged");
		expect(s.mergedCells).toEqual(["A3:B4"]);
	});

	it("reads sheet visibility and an empty sheet", async () => {
		const wb = await openOds(await loadFixture("ods-edge.ods"));
		expect(wb.sheets.map((s) => [s.name, s.state, s.visible])).toEqual([
			["Data", "visible", true],
			["Hidden", "hidden", false],
			["Blank", "visible", true],
		]);
		expect(wb.sheet("Hidden").cell("A1").value).toBe("secret");
		expect(wb.sheet("Blank").dimension).toBeUndefined();
	});
});

describe("openOds — adversarial-review regressions", () => {
	it("reads an ODF year 0–99 as the literal year, not the 1900s (Date.UTC two-digit trap)", () => {
		const [sheet] = parseOdsContent(
			odsDoc(
				'<table:table table:name="S"><table:table-row>' +
					'<table:table-cell office:value-type="date" office:date-value="0050-01-15"><text:p>x</text:p></table:table-cell>' +
					'<table:table-cell office:value-type="date" office:date-value="0001-06-30T13:30:00"><text:p>y</text:p></table:table-cell>' +
					"</table:table-row></table:table>",
			),
		);
		const a1 = sheet?.cells.get("A1")?.value as Date;
		const b1 = sheet?.cells.get("B1")?.value as Date;
		expect([a1.getUTCFullYear(), a1.getUTCMonth(), a1.getUTCDate()]).toEqual([50, 0, 15]);
		expect([b1.getUTCFullYear(), b1.getUTCHours(), b1.getUTCMinutes()]).toEqual([1, 13, 30]);
	});

	it("bounds materialized cells DOCUMENT-WIDE across repeat-bomb sheets, not per-sheet", () => {
		// Each sheet alone can hit the cap (200 rows × 16 384 cols > 2M); per-sheet counting would
		// let four sheets materialize ~8M cells. A shared budget caps the total at ~2M.
		const bomb =
			'<table:table table:name="B"><table:table-row table:number-rows-repeated="200">' +
			'<table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="16384"/>' +
			"</table:table-row></table:table>";
		const sheets = parseOdsContent(odsDoc(bomb + bomb + bomb + bomb));
		const total = sheets.reduce((n, s) => n + s.cells.size, 0);
		expect(total).toBeGreaterThan(0);
		expect(total).toBeLessThanOrEqual(2_000_000);
	}, 20_000);

	it("drops a span that the grid-edge clamp collapses to a single cell", () => {
		// A cell at the last column spanning 2 more columns clamps to XFD..XFD — not a real merge.
		const [sheet] = parseOdsContent(
			odsDoc(
				'<table:table table:name="S"><table:table-row>' +
					'<table:table-cell table:number-columns-repeated="16383"/>' +
					'<table:table-cell office:value-type="string" table:number-columns-spanned="3"><text:p>edge</text:p></table:table-cell>' +
					"</table:table-row></table:table>",
			),
		);
		expect(sheet?.cells.get("XFD1")?.value).toBe("edge");
		expect(sheet?.merges).toEqual([]);
	});
});

describe("openOds — unsupported accessors degrade (never throw)", () => {
	it("returns empty/undefined for the features .ods does not carry", async () => {
		const s = (await openOds(await loadFixture("odf-basic.ods"))).sheet("Sheet1");
		expect(s.style("A1")).toBeUndefined();
		expect(s.numberFormat("A1")).toBeUndefined();
		expect(s.formula("A1")).toBeUndefined();
		expect(s.comments).toEqual([]);
		expect(s.columns).toEqual([]);
		expect(s.rowProperties.size).toBe(0);
		expect(s.freeze).toBeUndefined();
		expect(await s.images()).toEqual([]);
	});
});

describe("openOds — typed failures", () => {
	it("throws not-a-zip on non-ZIP bytes", async () => {
		await expect(openOds(new Uint8Array([1, 2, 3, 4]))).rejects.toMatchObject({
			name: "XlsxError",
			code: "not-a-zip",
		});
	});

	it("refuses an encrypted document", async () => {
		await expect(openOds(await loadFixture("ods-encrypted.ods"))).rejects.toMatchObject({
			code: "unsupported",
		});
	});

	it("refuses a non-spreadsheet ODF (text document)", async () => {
		await expect(openOds(await loadFixture("ods-not-spreadsheet.ods"))).rejects.toMatchObject({
			code: "unsupported",
		});
	});

	it("fails typed when content.xml is missing", async () => {
		await expect(openOds(await loadFixture("ods-no-content.ods"))).rejects.toMatchObject({
			code: "missing-part",
		});
	});

	it("surfaces every failure as an XlsxError", async () => {
		const err = await openOds(new Uint8Array([0])).catch((e) => e);
		expect(err).toBeInstanceOf(XlsxError);
	});
});

describe("openOds — bridges to the xlsx writer (read → convert → write)", () => {
	it("converts an .ods to .xlsx through workbookToInput, values intact", async () => {
		const wb = await openOds(await loadFixture("odf-basic.ods"));
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(wb)));
		const s = rewritten.sheet("Sheet1");
		expect(s.cell("A1").value).toBe("hello");
		expect(s.cell("B1").value).toBe(42);
		expect(s.cell("F1").type).toBe("date");
		expect(s.cell("A2").value).toBe(84);
		expect(s.mergedCells).toContain("A4:B5");
		expect(rewritten.sheets.map((x) => x.name)).toEqual(["Sheet1", "Sheet2"]);
	});
});
