import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import type { ConditionalFormatting } from "../../types";
import { createMediaRegistry } from "../images";
import { createTableContext, streamWorksheet, worksheetXml } from "../sheet";
import { createStyleRegistry } from "../styles";

// F11.55: conditional-formatting emission is linear in rules plus blocks. Each block used to rescan
// every rule of the sheet, so n one-rule blocks cost n² (the bridge writes one block per range a
// producer declared, and the reader does not cap the count).

const DOUBLE_RULE = "</conditionalFormatting><conditionalFormatting";

describe("conditional-formatting emission cost", () => {
	it.each([
		[
			"the buffered writer",
			(cfs: readonly ConditionalFormatting[]) => () =>
				worksheetXml(
					{ name: "S", conditionalFormatting: cfs },
					[],
					0,
					false,
					createStyleRegistry(),
					createMediaRegistry(),
					createTableContext(),
					"S",
				),
		],
		[
			"the streaming writer",
			(cfs: readonly ConditionalFormatting[]) => () =>
				streamWorksheet(
					{ name: "S", conditionalFormatting: cfs },
					[],
					0,
					false,
					createStyleRegistry(),
					createMediaRegistry(),
					createTableContext(),
					"S",
				),
		],
	])(
		"is linear in one-rule blocks in %s",
		(_writer, render) => {
			const ratio = bestDoublingRatio(
				(n) =>
					render(
						Array.from({ length: n }, (_, i) => ({
							sqref: [`A${i + 1}`],
							rules: [
								{
									type: "cellIs",
									priority: 1,
									operator: "greaterThan",
									formulas: ["1"],
								},
							],
						})),
					),
				10_000,
			);
			expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
		},
		120_000,
	);

	it("emits each block's rules in its own element, numbered across the sheet", () => {
		const xml = worksheetXml(
			{
				name: "S",
				conditionalFormatting: [
					{
						sqref: ["A1"],
						rules: [
							{
								type: "cellIs",
								priority: 3,
								operator: "greaterThan",
								formulas: ["1"],
							},
							{ type: "cellIs", priority: 1, operator: "lessThan", formulas: ["0"] },
						],
					},
					{
						sqref: ["B1:B2"],
						rules: [
							{ type: "cellIs", priority: 2, operator: "equal", formulas: ["5"] },
						],
					},
				],
			},
			[],
			0,
			false,
			createStyleRegistry(),
			createMediaRegistry(),
			createTableContext(),
			"S",
		).xml;
		expect(xml).toContain(
			'<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="3" operator="greaterThan"><formula>1</formula></cfRule><cfRule type="cellIs" priority="1" operator="lessThan"><formula>0</formula></cfRule>' +
				DOUBLE_RULE +
				' sqref="B1:B2"><cfRule type="cellIs" priority="2" operator="equal"><formula>5</formula></cfRule></conditionalFormatting>',
		);
	});
});
