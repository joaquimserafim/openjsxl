import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

const FIXTURE = "excel-online-hyperlink-location-saveback.xlsx";

// Keep one logical contract for the Excel Online return and our rewritten copy. For example,
// Excel's `_x005f_x0041_` wire spelling and our `_x005F_x0041_` both expose `_x0041_` here.
function expectLogicalState(workbook: Workbook): void {
	expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Control", "_x0041_"]);
	expect(workbook.sheet("Control").cell("B12").value).toBe("AFTER");
	expect(workbook.sheet("_x0041_").cell("B4").value).toBe(
		"CORRECT TARGET — escape-shaped sheet / B4",
	);
	expect(
		workbook.sheet("Control").hyperlinks.map(({ ref, target, location }) => ({
			ref,
			target,
			location,
		})),
	).toEqual([
		{
			ref: "B5",
			location: "'_x0041_'!B4",
			target: undefined,
		},
		{
			ref: "B7",
			target: "https://example.com/",
			location: "_x0041_",
		},
	]);
}

describe("reader — Excel Online hyperlink location save-back (E20)", () => {
	it("pins the genuine producer package and protected location wire forms", async () => {
		const bytes = await loadFixture(FIXTURE);
		const zip = openZip(bytes, { verifyChecksums: true });

		// Reading every entry makes the fixture itself a CRC-checked interoperability input, not just
		// a source for the handful of XML parts asserted below.
		await Promise.all([...zip.entries.keys()].map((name) => zip.read(name)));
		const decodePart = async (name: string): Promise<string> =>
			new TextDecoder().decode(await zip.read(name));

		const appXml = await decodePart("docProps/app.xml");
		const workbookXml = await decodePart("xl/workbook.xml");
		const workbookRelsXml = await decodePart("xl/_rels/workbook.xml.rels");
		const sheetXml = await decodePart("xl/worksheets/sheet1.xml");
		const sheetRelsXml = await decodePart("xl/worksheets/_rels/sheet1.xml.rels");
		const entryNames = [...zip.entries.keys()];

		expect(appXml).toContain("<Application>Microsoft Excel Online</Application>");
		expect(appXml).toContain("<AppVersion>16.0300</AppVersion>");
		expect(workbookXml).toContain('<sheet name="_x005f_x0041_"');
		expect(sheetXml).toContain("location=\"'_x005f_x0041_'!B4\"");
		expect(sheetXml).toContain('location="_x005f_x0041_"');
		expect(sheetRelsXml).toContain('Target="https://example.com/" TargetMode="External"');

		// The PRE workbook lost its usable internal destination in Excel. The POST return must not
		// replace that repair with workbook external-link parts or recovery output.
		expect(entryNames.some((name) => name.startsWith("xl/externalLinks/"))).toBe(false);
		expect(entryNames.some((name) => /recovery/i.test(name))).toBe(false);
		expect(workbookXml).not.toContain("<externalReferences");
		expect(workbookRelsXml).not.toContain("/externalLink");
	});

	it("decodes the Excel Online return and bridges it without hyperlink drift", async () => {
		const before = await openXlsx(await loadFixture(FIXTURE), { verifyChecksums: true });
		expectLogicalState(before);

		const rewrittenBytes = await writeXlsx(await workbookToInput(before));
		const after = await openXlsx(rewrittenBytes, { verifyChecksums: true });
		expectLogicalState(after);
		expect(after.sheet("Control").hyperlinks).toEqual(before.sheet("Control").hyperlinks);

		const rewrittenZip = openZip(rewrittenBytes, { verifyChecksums: true });
		const rewrittenWorkbookXml = new TextDecoder().decode(
			await rewrittenZip.read("xl/workbook.xml"),
		);
		const rewrittenSheetXml = new TextDecoder().decode(
			await rewrittenZip.read("xl/worksheets/sheet1.xml"),
		);

		// openjsxl emits the same protected logical values with its canonical uppercase hex spelling.
		expect(rewrittenWorkbookXml).toContain('<sheet name="_x005F_x0041_"');
		expect(rewrittenSheetXml).toContain("location=\"'_x005F_x0041_'!B4\"");
		expect(rewrittenSheetXml).toContain('location="_x005F_x0041_"');
		expect(
			[...rewrittenZip.entries.keys()].some((name) => name.startsWith("xl/externalLinks/")),
		).toBe(false);
	});
});
