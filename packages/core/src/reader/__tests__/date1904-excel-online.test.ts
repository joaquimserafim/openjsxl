import { createHash } from "node:crypto";
import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";

const FIXTURE = "excel-online-1904-saveback.xlsx";
// The workbook as openjsxl wrote it and the owner opened it, before Excel saved it back.
const SENT = "openjsxl-1904-check-sent.xlsx";

// F11.8 — the genuine Microsoft Excel Online return of the openjsxl-authored 1.1.0 workbook written
// with `date1904: true`. Excel kept the 1904 flag, showed each date as intended, and computed YEAR,
// MONTH, DAY and a `Bn=DATE(y,m,d)` check that is TRUE only under the 1904 system.
// See packages/fixtures/data/README.md.

// [row, intended date, 1904 serial]; YEAR/MONTH/DAY follow from the date.
const DATES: ReadonlyArray<readonly [number, string, number]> = [
	[2, "1904-01-01", 0],
	[3, "1904-02-29", 59],
	[4, "1999-12-31", 35063],
	[5, "2000-02-29", 35123],
	[6, "2024-01-02", 43831],
];

function expectLogicalState(workbook: Workbook): void {
	expect(workbook.date1904).toBe(true);
	const control = workbook.sheet("Control");
	for (const [row, iso, serial] of DATES) {
		const [year, month, day] = iso.split("-").map(Number);
		expect(control.cell(`A${row}`).value).toBe(iso);
		expect(control.cell(`B${row}`).value).toEqual(new Date(`${iso}T00:00:00.000Z`));
		expect(control.cell(`C${row}`).value).toBe(year);
		expect(control.cell(`D${row}`).value).toBe(month);
		expect(control.cell(`E${row}`).value).toBe(day);
		expect(control.cell(`F${row}`).value).toBe(serial); // the stored serial, format "0"
		expect(control.cell(`G${row}`).value).toBe(serial);
		expect(control.formula(`H${row}`)).toBe(`B${row}=DATE(${year},${month},${day})`);
		expect(control.cell(`H${row}`).value).toBe(true);
	}
	expect(control.cell("B12").value).toBe("AFTER"); // the owner's save-back marker
}

describe("reader — Excel Online 1904 save-back (F11.8)", () => {
	it("pins the genuine producer package, the kept 1904 flag and Excel's own date checks", async () => {
		const bytes = await loadFixture(FIXTURE);
		const zip = openZip(bytes, { verifyChecksums: true });
		await Promise.all([...zip.entries.keys()].map((name) => zip.read(name)));
		const part = async (name: string): Promise<string> =>
			new TextDecoder().decode(await zip.read(name));

		const appXml = await part("docProps/app.xml");
		const coreXml = await part("docProps/core.xml");
		const workbookXml = await part("xl/workbook.xml");
		const sheetXml = await part("xl/worksheets/sheet1.xml");
		const stylesXml = await part("xl/styles.xml");

		expect(appXml).toContain("<Application>Microsoft Excel Online</Application>");
		expect(appXml).toContain("<AppVersion>16.0300</AppVersion>");
		expect(workbookXml).toContain('rupBuild="30626"');
		expect(zip.entries.has("xl/calcChain.xml")).toBe(true);
		expect(workbookXml).toContain('<workbookPr date1904="1"');
		expect(stylesXml).toContain('formatCode="yyyy\\-mm\\-dd"'); // Excel escaped the hyphens
		// Excel computed every check itself (the sent workbook carried no cached values).
		expect(sheetXml).toContain('<c r="B6" s="2"><v>43831</v></c>');
		expect(sheetXml).toContain("<f>YEAR(B6)</f><v>2024</v>");
		expect(sheetXml).toContain("<f>MONTH(B6)</f><v>1</v>");
		expect(sheetXml).toContain("<f>DAY(B6)</f><v>2</v>");
		expect(sheetXml).toContain('<c r="H2" t="b"><f>B2=DATE(1904,1,1)</f><v>1</v></c>');
		expect(sheetXml).toContain('<c r="H6" t="b"><f>B6=DATE(2024,1,2)</f><v>1</v></c>');
		expect(coreXml).toContain("<dc:creator></dc:creator>");
		expect(coreXml).toContain("<cp:lastModifiedBy></cp:lastModifiedBy>");
		expect(workbookXml).not.toContain("absPath");
	});

	it("pins the sent workbook: openjsxl's bytes, with no cached check value and the marker unset", async () => {
		const bytes = await loadFixture(SENT);
		expect(createHash("sha256").update(bytes).digest("hex")).toBe(
			"8cd76ca7b0003fd5d4211aa0ced3df521a94f1559fba4cd0b5a25ab3c3f469f9",
		);
		const zip = openZip(bytes, { verifyChecksums: true });
		const decode = new TextDecoder();
		const sheetXml = decode.decode(await zip.read("xl/worksheets/sheet1.xml"));
		expect(decode.decode(await zip.read("xl/workbook.xml"))).toContain(
			'<workbookPr date1904="1"/>',
		);
		expect(sheetXml).toContain("<f>YEAR(B6)</f></c>");
		expect(sheetXml).not.toContain("</f><v>"); // every check value above is Excel's own
		expect((await openXlsx(bytes)).sheet("Control").cell("B12").value).toBe("BEFORE");
	});

	it("reads the Excel Online return on the 1904 system and bridges it without a date shift", async () => {
		const before = await openXlsx(await loadFixture(FIXTURE), { verifyChecksums: true });
		expectLogicalState(before);

		const rewrittenBytes = await writeXlsx(await workbookToInput(before));
		const after = await openXlsx(rewrittenBytes, { verifyChecksums: true });
		expectLogicalState(after);

		const rewrittenWorkbookXml = new TextDecoder().decode(
			await openZip(rewrittenBytes, { verifyChecksums: true }).read("xl/workbook.xml"),
		);
		expect(rewrittenWorkbookXml).toContain('<workbookPr date1904="1"/>');
	});
});
