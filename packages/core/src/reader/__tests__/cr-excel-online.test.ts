import { createHash } from "node:crypto";
import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

const FIXTURE = "excel-online-cr-saveback.xlsx";
// The workbook as openjsxl wrote it and the owner opened it, before Excel saved it back.
const SENT = "openjsxl-cr-check-sent.xlsx";

// F11.8 — the genuine Microsoft Excel Online return of the openjsxl-authored 1.1.0 CR check
// workbook. openjsxl writes a carriage return as `_x000D_` in every `ST_Xstring` field (D17,
// F11.45). Excel read each one as a real CR: its cached LEN and CODE checks are 3 and 13, against
// 9 and 95 for the literal `_x005F_x000D_` control. It wrote every field back as `_x000D_`, except
// the hyperlink location, which it lowercased to `_x000d_`. See packages/fixtures/data/README.md.

// The check cells' values as Excel computed them: [ref, value] for LEN and CODE(MID(…,2,1)).
const CHECKS: ReadonlyArray<readonly [string, number]> = [
	["C2", 3],
	["D2", 13],
	["C3", 3],
	["D3", 13],
	["C4", 3],
	["D4", 13],
	["C5", 9],
	["D5", 95],
];

// One logical contract for the Excel return and for our rewritten copy of it: decoded text, not
// wire spelling (Excel's lowercase and our uppercase hex both mean a CR).
function expectLogicalState(workbook: Workbook): void {
	const control = workbook.sheet("Control");
	for (const ref of ["B2", "B3", "B4"]) expect(control.cell(ref).value).toBe("a\rb");
	expect(control.formula("B3")).toBe('"a\rb"');
	expect(control.formula("B4")).toBe("CRText");
	expect(control.cell("B5").value).toBe("a_x000D_b"); // the literal control: nine characters
	for (const [ref, value] of CHECKS) expect(control.cell(ref).value).toBe(value);
	expect(workbook.definedNames).toEqual([{ name: "CRText", refersTo: '"a\rb"' }]);
	const cfFormulas = control.conditionalFormatting.map((block) =>
		block.rules.map((rule) => ("formulas" in rule ? rule.formulas : undefined)),
	);
	expect(cfFormulas).toEqual([[['LEN("a\rb")=3'], ['LEN("a\rb")=9']]]);
	expect(control.comments.map((c) => [c.ref, c.author, c.text])).toEqual([
		["B7", "openjsxl", "line 1\rline 2"],
		["B8", "Ada\rLovelace", "The author's name holds a CR."],
	]);
	expect(control.hyperlinks).toEqual([{ ref: "B9", location: "a\rb", display: "link" }]);
	expect(control.cell("B12").value).toBe("AFTER"); // the owner's save-back marker
}

describe("reader — Excel Online CR save-back (F11.8)", () => {
	it("pins the genuine producer package, its CR wire forms and Excel's own check results", async () => {
		const bytes = await loadFixture(FIXTURE);
		const zip = openZip(bytes, { verifyChecksums: true });
		// Reading every entry makes the fixture itself a CRC-checked interoperability input.
		await Promise.all([...zip.entries.keys()].map((name) => zip.read(name)));
		const part = async (name: string): Promise<string> =>
			new TextDecoder().decode(await zip.read(name));

		const appXml = await part("docProps/app.xml");
		const coreXml = await part("docProps/core.xml");
		const workbookXml = await part("xl/workbook.xml");
		const sheetXml = await part("xl/worksheets/sheet1.xml");
		const stringsXml = await part("xl/sharedStrings.xml");
		const commentsXml = await part("xl/comments1.xml");

		expect(appXml).toContain("<Application>Microsoft Excel Online</Application>");
		expect(appXml).toContain("<AppVersion>16.0300</AppVersion>");
		expect(workbookXml).toContain('rupBuild="30626"');
		expect(zip.entries.has("xl/calcChain.xml")).toBe(true);
		// Excel computed every check itself (the sent workbook carried no cached values).
		expect(sheetXml).toContain("<f>LEN(B2)</f><v>3</v>");
		expect(sheetXml).toContain("<f>CODE(MID(B2,2,1))</f><v>13</v>");
		expect(sheetXml).toContain("<f>CODE(MID(CRText,2,1))</f><v>13</v>");
		expect(sheetXml).toContain("<f>LEN(B5)</f><v>9</v>");
		expect(sheetXml).toContain("<f>CODE(MID(B5,2,1))</f><v>95</v>");
		// Excel's own wire forms for each CR field.
		expect(stringsXml).toContain("<t>a_x000D_b</t>");
		expect(stringsXml).toContain("<t>a_x005F_x000D_b</t>");
		expect(sheetXml).toContain('<c r="B3" t="str"><f>"a_x000D_b"</f><v>a_x000D_b</v></c>');
		expect(sheetXml).toContain('<c r="B4" t="str"><f>CRText</f><v>a_x000D_b</v></c>');
		expect(workbookXml).toContain('<definedName name="CRText">"a_x000D_b"</definedName>');
		expect(sheetXml).toContain('<formula>LEN("a_x000D_b")=3</formula>');
		expect(sheetXml).toContain('<formula>LEN("a_x000D_b")=9</formula>');
		expect(sheetXml).toContain('<hyperlink ref="B9" location="a_x000d_b" display="link"');
		expect(commentsXml).toContain("<author>Ada_x000D_Lovelace</author>");
		expect(commentsXml).toContain("<t>line 1_x000D_line 2</t>");
		// The return carries no personal data: empty creator and last editor, no local path.
		expect(coreXml).toContain("<dc:creator></dc:creator>");
		expect(coreXml).toContain("<cp:lastModifiedBy></cp:lastModifiedBy>");
		expect(workbookXml).not.toContain("absPath");
	});

	it("pins the sent workbook: openjsxl's bytes, with no cached check value and the marker unset", async () => {
		const bytes = await loadFixture(SENT);
		expect(createHash("sha256").update(bytes).digest("hex")).toBe(
			"8c8b048a4e10840677ed0620f50065929878d7c88e18cff979139c8a4fcac2eb",
		);
		const sheetXml = new TextDecoder().decode(
			await openZip(bytes, { verifyChecksums: true }).read("xl/worksheets/sheet1.xml"),
		);
		expect(sheetXml).toContain("<f>LEN(B2)</f></c>");
		expect(sheetXml).not.toContain("</f><v>"); // every check value above is Excel's own
		expect(sheetXml).toContain('<f>"a_x000D_b"</f>');
		expect((await openXlsx(bytes)).sheet("Control").cell("B12").value).toBe("BEFORE");
	});

	it("decodes the Excel Online return and bridges it without CR drift", async () => {
		const before = await openXlsx(await loadFixture(FIXTURE), { verifyChecksums: true });
		expectLogicalState(before);

		const rewrittenBytes = await writeXlsx(await workbookToInput(before));
		const after = await openXlsx(rewrittenBytes, { verifyChecksums: true });
		expectLogicalState(after);

		const rewrittenZip = openZip(rewrittenBytes, { verifyChecksums: true });
		const rewrittenSheetXml = new TextDecoder().decode(
			await rewrittenZip.read("xl/worksheets/sheet1.xml"),
		);
		// openjsxl re-emits the CR in its canonical uppercase spelling, the location included.
		expect(rewrittenSheetXml).toContain('<f>"a_x000D_b"</f>');
		expect(rewrittenSheetXml).toContain('location="a_x000D_b"');
		expect(rewrittenSheetXml).not.toContain("_x000d_");
	});
});
