import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import type { DxfStyle } from "../../types";
import { MAX_CF_FORMULAS, parseConditionalFormatting } from "../conditional-formatting";
import { parseDxfs } from "../dxf";

// F9.3 — differential styles + conditional-formatting parser units.

const styles = (body: string): string =>
	`<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</styleSheet>`;
const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>${body}</worksheet>`;

describe("parseDxfs — differential styles kept RAW", () => {
	it("reads font, fill, border, numFmt, and alignment", () => {
		const dxfs = parseDxfs(
			styles(
				'<dxfs count="1"><dxf>' +
					'<font><b/><color rgb="FF9C0006"/></font>' +
					'<numFmt numFmtId="164" formatCode="0.00"/>' +
					'<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>' +
					'<border><left style="thin"><color rgb="FF00B050"/></left></border>' +
					'<alignment horizontal="center"/>' +
					"</dxf></dxfs>",
			),
		);
		expect(dxfs).toEqual([
			{
				font: { bold: true, color: { rgb: "FF9C0006" } },
				numberFormat: "0.00",
				fill: { bgColor: { rgb: "FFFFC7CE" } }, // patternType ABSENT — kept raw (bgColor visible)
				border: { left: { style: "thin", color: { rgb: "FF00B050" } } },
				alignment: { horizontal: "center" },
			},
		]);
	});

	it("keeps a solid fill's fgColor+bgColor exactly as written (no color swap)", () => {
		const [dxf] = parseDxfs(
			styles(
				'<dxfs><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFAABBCC"/><bgColor rgb="FFDDEEFF"/></patternFill></fill></dxf></dxfs>',
			),
		);
		expect(dxf?.fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FFAABBCC" },
			bgColor: { rgb: "FFDDEEFF" },
		});
	});

	it("stays index-aligned: an empty <dxf/> is a real slot", () => {
		const dxfs = parseDxfs(styles("<dxfs><dxf/><dxf><font><i/></font></dxf></dxfs>"));
		expect(dxfs).toEqual([{}, { font: { italic: true } }]);
	});

	it("returns [] for a stylesheet with no <dxfs>", () => {
		expect(parseDxfs(styles("<fonts/>"))).toEqual([]);
	});
});

describe("parseConditionalFormatting — rules + dxf resolution", () => {
	const DXFS: DxfStyle[] = [{ fill: { bgColor: { rgb: "FFFFC7CE" } } }, { font: { bold: true } }];

	it("resolves a cellIs rule's dxfId to an inline DxfStyle and keeps the formula", () => {
		const [block] = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>',
			),
			DXFS,
		);
		expect(block).toEqual({
			sqref: ["A1:A10"],
			rules: [
				{
					type: "cellIs",
					priority: 1,
					operator: "greaterThan",
					dxf: { fill: { bgColor: { rgb: "FFFFC7CE" } } },
					formulas: ["5"],
				},
			],
		});
	});

	it("reads a color scale, a data bar, and an icon set", () => {
		const [block] = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="B1:B5">' +
					'<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF638EC6"/></colorScale></cfRule>' +
					'<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>' +
					'<cfRule type="iconSet" priority="3"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>' +
					"</conditionalFormatting>",
			),
			DXFS,
		);
		const rules = block?.rules ?? [];
		expect(rules[0]).toEqual({
			type: "colorScale",
			priority: 1,
			cfvo: [{ type: "min" }, { type: "max" }],
			colors: [{ rgb: "FFFFFFFF" }, { rgb: "FF638EC6" }],
		});
		expect(rules[1]).toEqual({
			type: "dataBar",
			priority: 2,
			cfvo: [{ type: "min" }, { type: "max" }],
			color: { rgb: "FF638EC6" },
		});
		expect(rules[2]).toEqual({
			type: "iconSet",
			priority: 3,
			iconSet: "3TrafficLights1",
			cfvo: [
				{ type: "percent", val: "0" },
				{ type: "percent", val: "33" },
				{ type: "percent", val: "67" },
			],
		});
	});

	it("SKIPS x14 conditional formatting under <extLst>, keeping only the main block", () => {
		const blocks = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="1" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>' +
					'<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
					'<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:dataBar/></x14:cfRule></x14:conditionalFormatting>' +
					"</x14:conditionalFormattings></ext></extLst>",
			),
			DXFS,
		);
		expect(blocks).toHaveLength(1);
		expect(blocks[0]?.rules[0]?.type).toBe("expression");
	});

	it("drops non-canonical sqref tokens and unknown rule types (degrade into the writable set)", () => {
		const blocks = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A:A B2:C3"><cfRule type="cellIs" priority="1" dxfId="0"><formula>1</formula></cfRule>' +
					'<cfRule type="madeUpType" priority="2"/></conditionalFormatting>',
			),
			DXFS,
		);
		expect(blocks[0]?.sqref).toEqual(["B2:C3"]); // whole-column A:A dropped
		expect(blocks[0]?.rules).toHaveLength(1); // unknown type dropped
	});

	// F9.3 review regressions: reader output must always be writer-acceptable (shared bounds).
	it('degrades an XML-unsafe cfvo type to "num" (never returns a type the writer rejects)', () => {
		const [block] = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>' +
					'<cfvo type="mi&#1;n"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF000000"/>' +
					"</colorScale></cfRule></conditionalFormatting>",
			),
			DXFS,
		);
		const rule = block?.rules[0];
		const cfvo = rule?.type === "colorScale" ? rule.cfvo : [];
		expect(cfvo[0]?.type).toBe("num"); // control-char type degraded, not surfaced verbatim
	});

	it("DROPS an out-of-count colorScale / dataBar / iconSet (schema-invalid → unrepresentable)", () => {
		// 1-cfvo colorScale, 1-cfvo dataBar, 2-cfvo 3TrafficLights1 — all dropped.
		const blocks = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1">' +
					'<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><color rgb="FFFFFFFF"/></colorScale></cfRule>' +
					'<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><color rgb="FF638EC6"/></dataBar></cfRule>' +
					'<cfRule type="iconSet" priority="3"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/></iconSet></cfRule>' +
					"</conditionalFormatting>",
			),
			DXFS,
		);
		expect(blocks).toEqual([]); // no valid rule left → block dropped
	});

	it("takes cf formulas in stored form (leading = stripped)", () => {
		const [block] = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>=A1&gt;0</formula></cfRule></conditionalFormatting>',
			),
			DXFS,
		);
		const rule = block?.rules[0];
		expect(rule?.type === "expression" ? rule.formulas : undefined).toEqual(["A1>0"]);
	});

	it("ignores the 4th+ formula of a malformed rule (shared MAX_CF_FORMULAS bound, F9.6)", () => {
		// CT_CfRule allows ≤ 3 <formula> children; a hostile/malformed rule carrying more must still
		// come back writer-legal (the writer rejects > MAX_CF_FORMULAS — shared bound).
		const [block] = parseConditionalFormatting(
			sheet(
				'<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">' +
					"<formula>1</formula><formula>2</formula><formula>3</formula><formula>4</formula><formula>5</formula>" +
					"</cfRule></conditionalFormatting>",
			),
			DXFS,
		);
		const rule = block?.rules[0];
		expect(rule?.type === "expression" ? rule.formulas : undefined).toEqual(["1", "2", "3"]);
		expect(MAX_CF_FORMULAS).toBe(3); // the schema's maxOccurs — pinned so both sides move together
	});
});

describe("conditional formatting — verbatim read of the openpyxl fixture", () => {
	it("reads ≥10 rule types with resolved dxfs from openpyxl-condformat.xlsx", async () => {
		const book = await openXlsx(await loadFixture("openpyxl-condformat.xlsx"));
		const blocks = book.sheet("CF").conditionalFormatting;
		const rules = blocks.flatMap((b) => b.rules);
		const types = new Set<string>(rules.map((r) => r.type));
		// The full base ST_CfType surface authored in the fixture.
		for (const t of [
			"cellIs",
			"expression",
			"colorScale",
			"dataBar",
			"iconSet",
			"top10",
			"aboveAverage",
			"containsText",
			"duplicateValues",
			"beginsWith",
		]) {
			expect(types.has(t)).toBe(true);
		}
		// Two red-fill rules share one producer dxf → deep-equal resolved DxfStyle.
		const withFill = rules.filter(
			(r) =>
				r.type !== "colorScale" &&
				r.type !== "dataBar" &&
				r.type !== "iconSet" &&
				"dxf" in r &&
				r.dxf?.fill,
		);
		expect(withFill.length).toBeGreaterThanOrEqual(2);
		// The dxf fill is raw (openpyxl solid: patternType + fg + bg) — NOT swapped.
		const firstFill = withFill[0];
		const fill = firstFill && "dxf" in firstFill ? firstFill.dxf?.fill : undefined;
		expect(fill?.patternType).toBe("solid");
	});
});
