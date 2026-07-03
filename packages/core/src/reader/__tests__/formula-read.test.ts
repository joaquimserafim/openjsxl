import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";
import { parseFormulas } from "../worksheet";

// F5.4 — reading formula text. parseFormulas scans <f> elements; the shared-formula translation
// itself is unit-tested in ooxml/__tests__/formula.test.ts against openpyxl, so here we pin the
// scan: plain/array verbatim, shared master + translated dependents, dataTable skipped, entities.

const sheet = (body: string) =>
	`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;

describe("parseFormulas — units", () => {
	it("reads a plain formula verbatim, keyed by ref", () => {
		const f = parseFormulas(sheet('<row r="1"><c r="A1"><f>B1+C1</f><v>3</v></c></row>'));
		expect(f.get("A1")).toBe("B1+C1");
	});

	it("decodes XML entities in the formula text", () => {
		const f = parseFormulas(sheet('<row r="1"><c r="A1"><f>A1&lt;B1</f></c></row>'));
		expect(f.get("A1")).toBe("A1<B1");
	});

	it("translates shared-formula dependents from the master", () => {
		const f = parseFormulas(
			sheet(
				'<row r="1"><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
					'<row r="2"><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
					'<row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>',
			),
		);
		expect(f.get("B1")).toBe("A1*2"); // master
		expect(f.get("B2")).toBe("A2*2"); // translated
		expect(f.get("B3")).toBe("A3*2");
	});

	it("resolves a dependent even when the master appears later (two-pass)", () => {
		const f = parseFormulas(
			sheet(
				'<row r="2"><c r="B2"><f t="shared" si="0"/></c></row>' +
					'<row r="1"><c r="B1"><f t="shared" ref="B1:B2" si="0">A1+1</f></c></row>',
			),
		);
		expect(f.get("B2")).toBe("A2+1");
	});

	it("returns an array-master formula verbatim", () => {
		const f = parseFormulas(
			sheet('<row r="1"><c r="D1"><f t="array" ref="D1:D3">A1:A3*2</f></c></row>'),
		);
		expect(f.get("D1")).toBe("A1:A3*2");
	});

	it("skips a dataTable formula (no reusable text) and an orphan shared dependent", () => {
		const f = parseFormulas(
			sheet(
				'<row r="1"><c r="E1"><f t="dataTable" ref="E1:E2" dt2D="0" dtr="0" r1="A1"/></c>' +
					'<c r="F1"><f t="shared" si="9"/></c></row>',
			),
		);
		expect(f.has("E1")).toBe(false);
		expect(f.has("F1")).toBe(false); // no master for si=9
	});

	it("translates a shared whole-column formula dependent (review regression)", () => {
		const f = parseFormulas(
			sheet(
				'<row r="1"><c r="B1"><f t="shared" ref="B1:C1" si="0">SUM(A:A)</f></c>' +
					'<c r="C1"><f t="shared" si="0"/></c></row>',
			),
		);
		expect(f.get("B1")).toBe("SUM(A:A)");
		expect(f.get("C1")).toBe("SUM(B:B)"); // whole-column shifted, not left unchanged
	});

	it("ignores a stray <c><f> outside <sheetData> (review regression)", () => {
		// A real value cell A1, plus a stray formula-bearing <c r="A1"> in an oleObjects block; the
		// stray must NOT fabricate a formula on the real cell.
		const xml =
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
			'<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>' +
			'<oleObjects><c r="A1"><f>SUM(Z1:Z9)</f></c></oleObjects></worksheet>';
		expect(parseFormulas(xml).has("A1")).toBe(false);
	});

	it("degrades dependents of an over-long shared master (review regression: O(n·m) guard)", () => {
		const huge = `${"A1+".repeat(3000)}A1`; // 9001 chars, past MAX_FORMULA_LEN
		const f = parseFormulas(
			sheet(
				`<row r="1"><c r="B1"><f t="shared" ref="B1:B2" si="0">${huge}</f></c></row>` +
					'<row r="2"><c r="B2"><f t="shared" si="0"/></c></row>',
			),
		);
		expect(f.get("B1")).toBe(huge); // the master keeps its own (over-long) text
		expect(f.has("B2")).toBe(false); // dependent degrades rather than driving quadratic work
	});
});

describe("formula(ref) — real fixtures (e2e)", () => {
	it("reads basic.xlsx's cached formula as live text", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"));
		const s = wb.sheet(wb.sheets[0]!.name);
		expect(s.formula("E1")).toBe("B1*2");
		expect(s.cell("E1").value).toBe(84); // cached result still available
		expect(s.formula("A1")).toBeUndefined(); // a value cell has no formula
	});

	it("reads and translates the shared-formula fixture", async () => {
		const s = (await openXlsx(await loadFixture("shared-formula.xlsx"))).sheet("Calc");
		expect(s.formula("B1")).toBe("A1*2");
		expect(s.formula("B2")).toBe("A2*2"); // translated dependent
		expect(s.formula("B3")).toBe("A3*2");
		expect(s.formula("D1")).toBe("A1:A3*2"); // array master
		expect(s.cell("B2").value).toBe(40); // cached result
	});

	it("reads a formula alongside an error cached value", async () => {
		const s = (await openXlsx(await loadFixture("errors.xlsx"))).sheet(
			(await openXlsx(await loadFixture("errors.xlsx"))).sheets[0]!.name,
		);
		expect(s.formula("A1")).toBe("5/0");
		expect(s.cell("A1").type).toBe("error");
		expect(s.cell("A1").value).toBe("#DIV/0!");
	});
});
