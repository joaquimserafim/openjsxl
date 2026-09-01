import { describe, expect, it } from "vitest";
import { parseConditionalFormatting } from "../conditional-formatting";
import { MAX_DV_TEXT_LEN, parseDataValidations } from "../data-validation";
import { parseTable } from "../table";

// F11.20 — the three other ST_Formula parsers (data validation, conditional formatting, table
// formulas) decode the COMPLETE element text as ST_Xstring exactly once, BEFORE the leading-`=`
// strip, the XML-safety gate, the inline-list ceiling, and the formula-count cap — so every value
// the model carries is logical, and a decoded value the writer would refuse degrades away.

const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>${body}</worksheet>`;
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

describe("parseDataValidations — formula1/formula2 are ST_Xstring (F11.20)", () => {
	const dv = (f1: string, extra = "") =>
		parseDataValidations(
			sheet(
				`<dataValidations count="1"><dataValidation type="custom" sqref="A1"${extra}><formula1>${f1}</formula1></dataValidation></dataValidations>`,
			),
		)[0];

	it("decodes raw and protected escapes, and strips an ENCODED leading =", () => {
		expect(dv("'_x0041_'!B1")?.formula1).toBe("'A'!B1");
		expect(dv("'_x005F_x0041_'!B1")?.formula1).toBe("'_x0041_'!B1");
		expect(dv("_x003D_A1_x003E_1")?.formula1).toBe("A1>1"); // `=A1>1` after decoding
	});

	it("measures the inline-list ceiling on LOGICAL text", () => {
		// 200 escapes inside quotes: logical 202 chars (≤ MAX_DV_TEXT_LEN), wire 1,402 chars.
		const wire = `"${"_x0041_".repeat(200)}"`;
		expect(wire.length).toBeGreaterThan(MAX_DV_TEXT_LEN);
		const [rule] = parseDataValidations(
			sheet(
				`<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>${wire}</formula1></dataValidation></dataValidations>`,
			),
		);
		expect(rule?.formula1).toBe(`"${"A".repeat(200)}"`);
	});

	it("drops an operand whose DECODED text is XML-unsafe; a split escape still decodes", () => {
		expect(dv("A1_x0001_")?.formula1).toBeUndefined();
		expect(dv("_x00&#52;1_")?.formula1).toBe("A"); // entity split
		expect(dv("<![CDATA[_x00]]>41_")?.formula1).toBe("A"); // CDATA split
	});
});

describe("parseConditionalFormatting — <formula> is ST_Xstring (F11.20)", () => {
	// The rule union carries `formulas` only on highlight rules — narrow with `in`, never a cast.
	const formulasOf = (wire: string[]): readonly string[] | undefined => {
		const rule = parseConditionalFormatting(
			sheet(
				`<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1">${wire
					.map((f) => `<formula>${f}</formula>`)
					.join("")}</cfRule></conditionalFormatting>`,
			),
			undefined,
		)[0]?.rules[0];
		return rule !== undefined && "formulas" in rule ? rule.formulas : undefined;
	};

	it("decodes raw and protected escapes, and strips an ENCODED leading =", () => {
		expect(formulasOf(["'_x0041_'!B1"])).toEqual(["'A'!B1"]);
		expect(formulasOf(["'_x005F_x0041_'!B1"])).toEqual(["'_x0041_'!B1"]);
		expect(formulasOf(["_x003D_A1_x003E_1"])).toEqual(["A1>1"]);
	});

	it("drops an unsafe decoded operand and still caps at MAX_CF_FORMULAS on the decoded list", () => {
		expect(formulasOf(["A1_x0001_"])).toBeUndefined();
		expect(formulasOf(["_x0031_", "_x0032_", "_x0033_", "_x0034_"])).toEqual(["1", "2", "3"]);
	});

	it("decodes the COMPLETE text — an escape split by an entity or CDATA still decodes", () => {
		expect(formulasOf(["_x00&#52;1_"])).toEqual(["A"]);
		expect(formulasOf(["<![CDATA[_x00]]>41_"])).toEqual(["A"]);
	});
});

describe("parseTable — totalsRowFormula / calculatedColumnFormula are ST_Xstring (F11.20)", () => {
	const table = (totals: string, calc: string) =>
		parseTable(
			`<table ${NS} name="T" displayName="T" ref="A1:B3"><tableColumns count="2">` +
				'<tableColumn id="1" name="Item"/>' +
				`<tableColumn id="2" name="Amount"><totalsRowFormula>${totals}</totalsRowFormula><calculatedColumnFormula>${calc}</calculatedColumnFormula></tableColumn>` +
				"</tableColumns></table>",
		)?.columns[1];

	it("decodes both formulas once (raw → character, protected → literal)", () => {
		const c = table("SUBTOTAL(109,[_x005F_x0041_])", '[Item]&amp;"_x0041_"');
		expect(c?.totalsRowFormula).toBe("SUBTOTAL(109,[_x0041_])");
		expect(c?.calculatedColumnFormula).toBe('[Item]&"A"');
	});

	it("omits a formula whose DECODED text the writer would refuse — the column itself carries", () => {
		const c = table("SUM([Amount])_x0001_", "[Item]_xD800_");
		expect(c?.name).toBe("Amount");
		expect(c?.totalsRowFormula).toBeUndefined();
		expect(c?.calculatedColumnFormula).toBeUndefined();
	});

	it("decodes the COMPLETE text — a split escape still decodes", () => {
		expect(table("_x00&#52;1_", "x")?.totalsRowFormula).toBe("A"); // entity split
		expect(table("x", "<![CDATA[_x00]]>41_")?.calculatedColumnFormula).toBe("A"); // CDATA split
	});
});
