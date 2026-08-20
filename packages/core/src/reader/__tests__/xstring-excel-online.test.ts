import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

const FIXTURE = "excel-online-xstring-saveback.xlsx";

// Keep one logical contract for the genuine Excel Online file and our rewritten copy. This makes
// the regression prove both producer compatibility and bridge fidelity without duplicating asserts.
function expectLogicalState(workbook: Workbook): void {
	expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Control", "_x0041_"]);
	expect(workbook.sheet("Control").cell("B12").value).toBe("AFTER");
	expect(workbook.sheet("_x0041_").cell("B4").value).toBe("_x0041_");
	expect(workbook.sheet("_x0041_").cell("B5").value).toBe("_x005F_x0041_");
	expect(workbook.definedNames).toEqual([{ name: "_x0041_", refersTo: "_x0041_!B4" }]);
	expect(workbook.sheet("_x0041_").comments).toEqual([
		{
			ref: "B7",
			author: "_x0041_",
			text: "F11.4 legacy Note author check.",
		},
	]);
}

describe("reader — Excel Online ST_Xstring save-back (F11.4)", () => {
	it("pins the genuine producer wire forms and absence of the earlier external-link corruption", async () => {
		const bytes = await loadFixture(FIXTURE);
		const zip = openZip(bytes);
		const decodePart = async (name: string): Promise<string> =>
			new TextDecoder().decode(await zip.read(name));

		const workbookXml = await decodePart("xl/workbook.xml");
		const workbookRelsXml = await decodePart("xl/_rels/workbook.xml.rels");
		const sheetRelsXml = await decodePart("xl/worksheets/_rels/sheet2.xml.rels");
		const commentsXml = await decodePart("xl/comments1.xml");
		const vmlXml = await decodePart("xl/drawings/vmlDrawing1.vml");
		const appXml = await decodePart("docProps/app.xml");

		// These are Excel Online's exact successful save-back bytes, including its lowercase `f` in
		// attribute escapes and uppercase `F` in element-text escapes.
		expect(appXml).toContain("<Application>Microsoft Excel Online</Application>");
		expect(appXml).toContain("<AppVersion>16.0300</AppVersion>");
		expect(workbookXml).toContain('<sheet name="_x005f_x0041_"');
		expect(workbookXml).toContain(
			'<definedName name="_x005f_x0041_">_x005F_x0041_!B4</definedName>',
		);
		expect(commentsXml).toContain("<author>_x005F_x0041_</author>");

		// The prior failed gate fabricated `[1]A!B4` plus external-link parts. Their absence is the
		// load-bearing interoperability result, while the VML checks keep the Note anchored at B7.
		expect([...zip.entries.keys()].some((name) => name.startsWith("xl/externalLinks/"))).toBe(
			false,
		);
		expect(workbookXml).not.toContain("<externalReferences");
		expect(workbookXml).not.toContain("[1]A!B4");
		expect(workbookRelsXml).not.toContain("/externalLink");
		expect(sheetRelsXml).toContain('/comments" Target="../comments1.xml"');
		expect(sheetRelsXml).toContain('/vmlDrawing" Target="../drawings/vmlDrawing1.vml"');
		expect(vmlXml).toContain("<x:Row>6</x:Row><x:Column>1</x:Column>");
	});

	it("decodes the Excel Online file and bridges it without logical drift", async () => {
		const before = await openXlsx(await loadFixture(FIXTURE), { verifyChecksums: true });
		expectLogicalState(before);

		const after = await openXlsx(await writeXlsx(await workbookToInput(before)), {
			verifyChecksums: true,
		});
		expectLogicalState(after);
	});
});
