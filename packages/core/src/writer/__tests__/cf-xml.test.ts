import { describe, expect, it } from "vitest";
import { conditionalFormattingXml } from "../sheet";
import { createStyleRegistry } from "../styles";

// conditionalFormattingXml's exits (F11.55 exported it for its own suite).

describe("conditionalFormattingXml", () => {
	it("emits nothing without blocks", () => {
		expect(conditionalFormattingXml("S", undefined, createStyleRegistry())).toBe("");
		expect(conditionalFormattingXml("S", [], createStyleRegistry())).toBe("");
	});

	it("emits each block with its own rules, priorities renumbered across the sheet", () => {
		expect(
			conditionalFormattingXml(
				"S",
				[
					{
						sqref: ["A1"],
						rules: [
							{
								type: "cellIs",
								priority: 5,
								operator: "greaterThan",
								formulas: ["1"],
							},
						],
					},
					{
						sqref: ["B1", "C1:C2"],
						rules: [
							{
								type: "cellIs",
								priority: 1,
								operator: "greaterThan",
								formulas: ["2"],
							},
							{
								type: "cellIs",
								priority: 5,
								operator: "greaterThan",
								formulas: ["3"],
							},
						],
					},
				],
				createStyleRegistry(),
			),
		).toBe(
			'<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>' +
				'<conditionalFormatting sqref="B1 C1:C2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule><cfRule type="cellIs" priority="3" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
		);
	});

	it("rejects a non-array and a block without rules", () => {
		// A JS caller can pass a value the types forbid.
		expect(() => conditionalFormattingXml("S", {} as never, createStyleRegistry())).toThrow(
			expect.objectContaining({
				code: "invalid-input",
				message: expect.stringContaining("conditionalFormatting must be an array"),
			}),
		);
		expect(() =>
			conditionalFormattingXml("S", [{ sqref: ["A1"], rules: [] }], createStyleRegistry()),
		).toThrow(/conditionalFormatting\[0\]\.rules must have at least one rule/);
	});

	// The blocks emitted are the blocks validated: a length that grows after validation adds none
	// (on HEAD it added an empty, schema-invalid block).
	it("emits only the blocks it validated when the array's length changes afterwards", () => {
		const blocks = [1, 2, 3].map((n) => ({
			sqref: [`A${n}`],
			rules: [
				{ type: "cellIs", priority: n, operator: "greaterThan", formulas: [String(n)] },
			],
		}));
		let lengthReads = 0;
		// Validation reads the length 4 times for 2 blocks (the empty check, then each loop test).
		const growing = new Proxy(blocks, {
			get: (target, key, receiver) =>
				key === "length"
					? ++lengthReads <= 4
						? 2
						: 3
					: Reflect.get(target, key, receiver),
		});
		const xml = conditionalFormattingXml("S", growing as never, createStyleRegistry());
		expect(xml.split("<conditionalFormatting ").length - 1).toBe(2);
		expect(xml).not.toContain('sqref=""');
	});
});
