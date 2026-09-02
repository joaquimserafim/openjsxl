import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { writeZip } from "../../writer/zip";
import { openXlsx } from "../workbook";
import { parseFormulas } from "../worksheet";

// F11.9 — shared-formula translation through the real reader surface, not only the helper: a shared
// master carrying a structured reference (opaque) and a lowercase cell reference (shifts, canonical
// uppercase) is expanded for its dependents by parseFormulas, surfaces through Worksheet.formula,
// and crosses the bridge (workbookToInput → writeXlsx) as plain per-cell formula text.

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const enc = new TextEncoder();

// Master at C1 covers C1:D3; dependents at C2 (Δ1,0) and D3 (Δ2,1) — the two openpyxl oracle offsets.
const MASTER = "SUM(Table1[A1])+a1+Table1[[#Data],[B2]]";
const SHEET_DATA =
	`<row r="1"><c r="C1"><f t="shared" ref="C1:D3" si="0">${MASTER}</f><v>0</v></c></row>` +
	'<row r="2"><c r="C2"><f t="shared" si="0"/><v>0</v></c></row>' +
	'<row r="3"><c r="D3"><f t="shared" si="0"/><v>0</v></c></row>';
const sheetXml = `<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData>${SHEET_DATA}</sheetData></worksheet>`;

function minimalXlsx(): Promise<Uint8Array> {
	const workbook =
		`<?xml version="1.0"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets>` +
		`<sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
	const wbRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
	const rootRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	const types =
		`<?xml version="1.0"?><Types xmlns="${CT}">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
		`</Types>`;
	return writeZip([
		{ name: "[Content_Types].xml", data: enc.encode(types) },
		{ name: "_rels/.rels", data: enc.encode(rootRels) },
		{ name: "xl/workbook.xml", data: enc.encode(workbook) },
		{ name: "xl/_rels/workbook.xml.rels", data: enc.encode(wbRels) },
		{ name: "xl/worksheets/sheet1.xml", data: enc.encode(sheetXml) },
	]);
}

const EXPECTED = {
	C1: MASTER, // the master's own text, verbatim
	C2: "SUM(Table1[A1])+A2+Table1[[#Data],[B2]]", // Δ(1,0): only `a1` moves, canonical uppercase
	D3: "SUM(Table1[A1])+B3+Table1[[#Data],[B2]]", // Δ(2,1)
};

describe("shared formulas with structured + lowercase references (F11.9)", () => {
	it("parseFormulas expands the dependents from the decoded master", () => {
		const f = parseFormulas(sheetXml);
		expect(Object.fromEntries(f)).toEqual(EXPECTED);
	});

	it("Worksheet.formula surfaces the translated text and the bridge carries it as plain formulas", async () => {
		const wb = await openXlsx(await minimalXlsx());
		const s = wb.sheet("S");
		expect(s.formula("C1")).toBe(EXPECTED.C1);
		expect(s.formula("C2")).toBe(EXPECTED.C2);
		expect(s.formula("D3")).toBe(EXPECTED.D3);

		// The bridge writes each dependent as its own plain formula (a documented degradation of the
		// shared grouping); the text — structured refs verbatim, `a1` already canonical — is unchanged.
		const again = (await openXlsx(await writeXlsx(await workbookToInput(wb)))).sheet("S");
		expect(again.formula("C1")).toBe(EXPECTED.C1);
		expect(again.formula("C2")).toBe(EXPECTED.C2);
		expect(again.formula("D3")).toBe(EXPECTED.D3);
	});

	it("degrades a malformed master (unbalanced bracket) without an untyped throw", () => {
		const f = parseFormulas(
			`<worksheet xmlns="${NS_MAIN}"><sheetData>` +
				'<row r="1"><c r="C1"><f t="shared" ref="C1:C2" si="0">a1+Table1[Col</f></c></row>' +
				'<row r="2"><c r="C2"><f t="shared" si="0"/></c></row>' +
				"</sheetData></worksheet>",
		);
		expect(f.get("C1")).toBe("a1+Table1[Col"); // master verbatim
		expect(f.get("C2")).toBe("A2+Table1[Col"); // the reference before the bracket still shifts
	});
});
