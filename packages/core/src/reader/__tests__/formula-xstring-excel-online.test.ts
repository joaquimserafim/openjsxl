import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

const FIXTURE = "excel-online-formula-xstring-saveback.xlsx";

// F11.20 — the genuine Microsoft Excel Online return of the openjsxl-authored POST workbook: every
// formula field references (or quotes) a worksheet literally named `_x0041_`. Excel kept the
// protected `ST_Xstring` spellings — its own lowercase `_x005f_x0041_` on the tab, our uppercase
// `_x005F_x0041_` inside the formulas — dropped the now-unnecessary quotes around the sheet
// reference, normalized the table formula to `T[[#This Row],[Item]]`, and moved the two
// cross-sheet DV/CF rules into `x14` extensions (a documented openjsxl drop). The PRE control,
// written with formulas raw, broke in the real UI: Excel decoded the reference to a sheet named
// `A` and fabricated an external workbook link. See packages/fixtures/data/README.md.

// One logical contract for the Excel return and for our rewritten copy of it: decoded text, not
// wire spelling (Excel's lowercase and our uppercase hex both mean the literal `_x0041_`).
function expectLogicalState(workbook: Workbook): void {
	expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Control", "_x0041_"]);
	const control = workbook.sheet("Control");
	expect(control.formula("A1")).toBe("_x0041_!B4*2"); // Excel dropped the quotes; same sheet
	expect(control.cell("A1").value).toBe(42);
	expect(control.formula("A2")).toBe('"_x0041_"'); // the literal survived as seven characters
	expect(control.cell("A2").value).toBe("_x0041_");
	expect(control.formula("A3")).toBe("LEN(A2)");
	expect(control.cell("A3").value).toBe(7);
	expect(control.formula("B6")).toBe('T[[#This Row],[Item]]&"_x0041_"');
	expect(control.cell("B6").value).toBe("a_x0041_");
	expect(control.cell("B12").value).toBe("AFTER"); // the owner's save-back marker
	expect(control.tables.map((t) => t.columns.map((c) => c.calculatedColumnFormula))).toEqual([
		[undefined, 'T[[#This Row],[Item]]&"_x0041_"'],
	]);
	// Excel stored the cross-sheet DV/CF rules only as x14 extensions, which openjsxl skips.
	expect(control.dataValidations).toEqual([]);
	expect(control.conditionalFormatting).toEqual([]);
	expect(workbook.sheet("_x0041_").cell("B4").value).toBe(21);
}

describe("reader — Excel Online formula ST_Xstring save-back (F11.20)", () => {
	it("pins the genuine producer package and the protected formula wire forms", async () => {
		const bytes = await loadFixture(FIXTURE);
		const zip = openZip(bytes, { verifyChecksums: true });
		// Reading every entry makes the fixture itself a CRC-checked interoperability input.
		await Promise.all([...zip.entries.keys()].map((name) => zip.read(name)));
		const part = async (name: string): Promise<string> =>
			new TextDecoder().decode(await zip.read(name));

		const appXml = await part("docProps/app.xml");
		const workbookXml = await part("xl/workbook.xml");
		const sheetXml = await part("xl/worksheets/sheet1.xml");
		const tableXml = await part("xl/tables/table1.xml");
		const entryNames = [...zip.entries.keys()];

		expect(appXml).toContain("<Application>Microsoft Excel Online</Application>");
		expect(appXml).toContain("<AppVersion>16.0300</AppVersion>");
		expect(workbookXml).toContain('<sheet name="_x005f_x0041_"');
		expect(sheetXml).toContain("<f>_x005F_x0041_!B4*2</f>");
		expect(sheetXml).toContain('<f>"_x005F_x0041_"</f>');
		expect(sheetXml).toContain('<f>T[[#This Row],[Item]]&amp;"_x005F_x0041_"</f>');
		expect(tableXml).toContain(
			'<calculatedColumnFormula>T[[#This Row],[Item]]&amp;"_x005F_x0041_"</calculatedColumnFormula>',
		);
		// The cross-sheet rules live in x14 — still carrying the protected text — not in the main part.
		expect(sheetXml).toContain("<xm:f>A1=_x005F_x0041_!B4*2</xm:f>");
		expect(sheetXml).toContain("<xm:f>C1=_x005F_x0041_!B4</xm:f>");
		expect(sheetXml).not.toContain("<dataValidations");
		expect(sheetXml).not.toContain("<conditionalFormatting ");

		// The PRE workbook was repaired into an external link; the POST return must carry none.
		expect(entryNames.some((name) => name.startsWith("xl/externalLinks/"))).toBe(false);
		expect(entryNames.some((name) => /recovery/i.test(name))).toBe(false);
		expect(workbookXml).not.toContain("<externalReferences");
	});

	it("decodes the Excel Online return and bridges it without formula drift", async () => {
		const before = await openXlsx(await loadFixture(FIXTURE), { verifyChecksums: true });
		expectLogicalState(before);

		const rewrittenBytes = await writeXlsx(await workbookToInput(before));
		const after = await openXlsx(rewrittenBytes, { verifyChecksums: true });
		expectLogicalState(after);

		const rewrittenZip = openZip(rewrittenBytes, { verifyChecksums: true });
		const decode = new TextDecoder();
		const rewrittenWorkbookXml = decode.decode(await rewrittenZip.read("xl/workbook.xml"));
		const rewrittenSheetXml = decode.decode(
			await rewrittenZip.read("xl/worksheets/sheet1.xml"),
		);
		const rewrittenTableXml = decode.decode(await rewrittenZip.read("xl/tables/table1.xml"));

		// openjsxl re-emits the same protected logical values in its canonical uppercase spelling.
		expect(rewrittenWorkbookXml).toContain('<sheet name="_x005F_x0041_"');
		expect(rewrittenSheetXml).toContain("<f>_x005F_x0041_!B4*2</f>");
		expect(rewrittenSheetXml).toContain('<f>"_x005F_x0041_"</f>');
		expect(rewrittenTableXml).toContain(
			'<calculatedColumnFormula>T[[#This Row],[Item]]&amp;"_x005F_x0041_"</calculatedColumnFormula>',
		);
		expect(
			[...rewrittenZip.entries.keys()].some((name) => name.startsWith("xl/externalLinks/")),
		).toBe(false);
	});
});
