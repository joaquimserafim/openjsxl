import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import {
	MAX_DV_TEXT_LEN,
	MAX_DV_TITLE_LEN,
	MAX_SQREF_RANGES,
	parseDataValidations,
} from "../data-validation";

// F9.2 — data-validation parser units. The parser is TOLERANT (never throws): unknown types/operators
// degrade, over-long prompt/error text clamps, a range-less rule drops, and worksheet-level x14
// validations under <extLst> are skipped (decision 4). Bounds (decision 5) are single-sourced here and
// shared with the writer.

// Wrap a <dataValidations> block in a minimal worksheet so the scan sees realistic surroundings.
const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>${body}</worksheet>`;

describe("parseDataValidations — types, operands, and flags", () => {
	it("reads a whole-number rule with operator, bounds, prompts, and errors", () => {
		const dvs = parseDataValidations(
			sheet(
				'<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1"' +
					' showInputMessage="1" showErrorMessage="1" errorStyle="stop" promptTitle="Q" prompt="1-100"' +
					' errorTitle="Bad" error="out of range" sqref="A2:A10">' +
					"<formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>",
			),
		);
		expect(dvs).toEqual([
			{
				sqref: ["A2:A10"],
				type: "whole",
				operator: "between",
				formula1: "1",
				formula2: "100",
				allowBlank: true,
				showInputMessage: true,
				showErrorMessage: true,
				errorStyle: "stop",
				promptTitle: "Q",
				prompt: "1-100",
				errorTitle: "Bad",
				error: "out of range",
			},
		]);
	});

	it("splits a multi-range sqref into symbolic tokens", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="decimal" sqref="B2:B10 D2:D10 F5">' +
					"<formula1>0</formula1></dataValidation></dataValidations>",
			),
		);
		expect(dv?.sqref).toEqual(["B2:B10", "D2:D10", "F5"]);
	});

	it("keeps the inline-list quotes and a cross-sheet range source verbatim", () => {
		const dvs = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="list" sqref="C1"><formula1>"a,b,c"</formula1></dataValidation>' +
					'<dataValidation type="list" sqref="D1"><formula1>Sheet2!$A$1:$A$3</formula1></dataValidation></dataValidations>',
			),
		);
		expect(dvs[0]?.formula1).toBe('"a,b,c"');
		expect(dvs[1]?.formula1).toBe("Sheet2!$A$1:$A$3");
	});

	it("inverts showDropDown: a file `1` (arrow hidden) reads as false, `0` as true", () => {
		const dvs = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="list" showDropDown="1" sqref="A1"><formula1>"x"</formula1></dataValidation>' +
					'<dataValidation type="list" showDropDown="0" sqref="B1"><formula1>"y"</formula1></dataValidation></dataValidations>',
			),
		);
		expect(dvs[0]?.showDropDown).toBe(false); // file 1 ⇒ hidden ⇒ intuitive false
		expect(dvs[1]?.showDropDown).toBe(true);
	});

	it("degrades an unknown type to none and drops an unknown operator/errorStyle", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="bogus" operator="nope" errorStyle="loud" sqref="A1"/></dataValidations>',
			),
		);
		expect(dv?.type).toBe("none");
		expect(dv?.operator).toBeUndefined();
		expect(dv?.errorStyle).toBeUndefined();
	});

	it("reads an input-message-only rule (no type attribute ⇒ none)", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation showInputMessage="1" promptTitle="Note" prompt="free-form" sqref="J2:J10"/></dataValidations>',
			),
		);
		expect(dv).toEqual({
			sqref: ["J2:J10"],
			type: "none",
			showInputMessage: true,
			promptTitle: "Note",
			prompt: "free-form",
		});
	});

	it("treats an empty <formula1> as absent", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="custom" sqref="A1"><formula1></formula1></dataValidation></dataValidations>',
			),
		);
		expect(dv?.formula1).toBeUndefined();
	});
});

describe("parseDataValidations — tolerant degrades (decision 4 & 5)", () => {
	it("SKIPS worksheet-level x14 validations under <extLst>, keeping only the main block", () => {
		const xml = sheet(
			'<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"main"</formula1></dataValidation></dataValidations>' +
				'<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
				'<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
				'<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Other!$A$1:$A$9</xm:f></x14:formula1>' +
				"<xm:sqref>B1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>",
		);
		const dvs = parseDataValidations(xml);
		expect(dvs).toHaveLength(1);
		expect(dvs[0]?.formula1).toBe('"main"');
		expect(dvs[0]?.sqref).toEqual(["A1"]);
	});

	it("drops a rule whose sqref covers nothing", () => {
		expect(
			parseDataValidations(
				sheet(
					'<dataValidations><dataValidation type="list" sqref=""><formula1>"x"</formula1></dataValidation></dataValidations>',
				),
			),
		).toEqual([]);
		expect(
			parseDataValidations(
				sheet(
					'<dataValidations><dataValidation type="list"><formula1>"x"</formula1></dataValidation></dataValidations>',
				),
			),
		).toEqual([]);
	});

	it("clamps an over-long prompt/error title and body to the shared bounds", () => {
		const longTitle = "T".repeat(MAX_DV_TITLE_LEN + 20);
		const longBody = "B".repeat(MAX_DV_TEXT_LEN + 50);
		const [dv] = parseDataValidations(
			sheet(
				`<dataValidations><dataValidation type="none" promptTitle="${longTitle}" prompt="${longBody}" sqref="A1"/></dataValidations>`,
			),
		);
		expect(dv?.promptTitle).toHaveLength(MAX_DV_TITLE_LEN);
		expect(dv?.prompt).toHaveLength(MAX_DV_TEXT_LEN);
	});

	it("caps the number of sqref ranges (repeat-bomb guard) and drops an over-long inline list", () => {
		const many = Array.from(
			{ length: MAX_SQREF_RANGES + 100 },
			(_, i) => `A${(i % 1000) + 1}`,
		).join(" ");
		const [dv] = parseDataValidations(
			sheet(
				`<dataValidations><dataValidation type="whole" sqref="${many}"/></dataValidations>`,
			),
		);
		expect(dv?.sqref.length).toBe(MAX_SQREF_RANGES);

		const bigList = `"${"a,".repeat(MAX_DV_TEXT_LEN)}"`;
		expect(
			parseDataValidations(
				sheet(
					`<dataValidations><dataValidation type="list" sqref="A1"><formula1>${bigList}</formula1></dataValidation></dataValidations>`,
				),
			),
		).toEqual([]);
	});

	it("returns [] for a sheet with no dataValidations block", () => {
		expect(parseDataValidations(sheet(""))).toEqual([]);
	});
});

// F9.2 adversarial-review regressions: every value the tolerant reader RETURNS must be one the strict
// writer ACCEPTS (shared bounds). The review found three reader→writer mismatches; these pin the fixes.
describe("parseDataValidations — reader degrades into the writer's accepted set (review regressions)", () => {
	it("DROPS non-canonical / out-of-grid sqref tokens, keeping only the writable ones", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="list" sqref="A:A B2:C3 $A$1 a1 A1:A2000000 Sheet2!A1 D5">' +
					'<formula1>"x"</formula1></dataValidation></dataValidations>',
			),
		);
		// Whole-column, absolute, lowercase, out-of-grid, and cross-sheet tokens are dropped.
		expect(dv?.sqref).toEqual(["B2:C3", "D5"]);
	});

	it("drops a rule whose sqref is ENTIRELY non-canonical (nothing writable left)", () => {
		expect(
			parseDataValidations(
				sheet(
					'<dataValidations><dataValidation type="list" sqref="A:A 1:1"><formula1>"x"</formula1></dataValidation></dataValidations>',
				),
			),
		).toEqual([]);
	});

	it("DROPS a prompt/error field carrying an XML-unsafe (decoded control) character", () => {
		// &#1; decodes to U+0001 — the writer's isXmlSafe would reject it, so the reader drops the field.
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="none" promptTitle="hi&#1;there" prompt="ok" sqref="A1"/></dataValidations>',
			),
		);
		expect(dv?.promptTitle).toBeUndefined(); // dropped
		expect(dv?.prompt).toBe("ok"); // the clean field survives
	});

	it("DROPS a formula operand carrying an XML-unsafe character", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="custom" sqref="A1"><formula1>IS&#7;BLANK(A1)</formula1></dataValidation></dataValidations>',
			),
		);
		expect(dv?.formula1).toBeUndefined();
	});

	it("takes formula operands in STORED form (a leading = is stripped, matching the writer)", () => {
		const [dv] = parseDataValidations(
			sheet(
				'<dataValidations><dataValidation type="custom" sqref="A1"><formula1>=A1&gt;0</formula1></dataValidation></dataValidations>',
			),
		);
		expect(dv?.formula1).toBe("A1>0"); // not "=A1>0"
	});
});

describe("data validation — verbatim read of the openpyxl fixture", () => {
	it("reads all 8 types + an input-message-only rule from openpyxl-datavalidation.xlsx", async () => {
		const book = await openXlsx(await loadFixture("openpyxl-datavalidation.xlsx"));
		const dvs = book.sheet("Rules").dataValidations;
		expect(dvs.map((d) => d.type)).toEqual([
			"whole",
			"decimal",
			"list",
			"list",
			"date",
			"time",
			"textLength",
			"custom",
			"none",
		]);
		// Multi-range sqref preserved.
		expect(dvs[1]?.sqref).toEqual(["B2:B10", "D2:D10"]);
		// The cross-sheet list source is carried verbatim, with the dropdown HIDDEN (showDropDown="1").
		expect(dvs[3]?.formula1).toBe("Lists!$A$1:$A$3");
		expect(dvs[3]?.showDropDown).toBe(false);
		// The inline list keeps its quotes; its dropdown is shown.
		expect(dvs[2]?.formula1).toBe('"Low,Medium,High"');
		expect(dvs[2]?.showDropDown).toBe(true);
		// The rich first rule carries its prompt + error text.
		expect(dvs[0]).toMatchObject({
			type: "whole",
			operator: "between",
			formula1: "1",
			formula2: "100",
			promptTitle: "Quantity",
			errorTitle: "Out of range",
			errorStyle: "stop",
		});
		// The other sheet carries no validations (degrade path unaffected).
		expect(book.sheet("Lists").dataValidations).toEqual([]);
	});
});
