import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import type { ConditionalFormatting } from "../../types";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { SheetInput } from "../types";
import { writeXlsx } from "../workbook";

const decoder = new TextDecoder();

const partText = async (bytes: Uint8Array, part: string): Promise<string> =>
	decoder.decode(await openZip(bytes).read(part));

async function roundTrip(cf: readonly ConditionalFormatting[]) {
	const sheet: SheetInput = { name: "S", rows: [[1], [2], [3]], conditionalFormatting: cf };
	const book = await openXlsx(await writeXlsx({ sheets: [sheet] }));
	return book.sheet("S").conditionalFormatting;
}

const rejects = (cf: readonly unknown[]) =>
	expect(
		writeXlsx({
			// biome-ignore lint/suspicious/noExplicitAny: hostile input past the types
			sheets: [{ name: "S", rows: [[1]], conditionalFormatting: cf as any }],
		}),
	).rejects.toThrow(XlsxError);

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		total += value.length;
	}
	const out = new Uint8Array(total);
	let at = 0;
	for (const c of chunks) {
		out.set(c, at);
		at += c.length;
	}
	return out;
}

describe("writeXlsx — conditional formatting round-trip", () => {
	it("round-trips highlight, colorScale, dataBar, and iconSet rules deep-equal", async () => {
		const cf: ConditionalFormatting[] = [
			{
				sqref: ["A1:A3"],
				rules: [
					{
						type: "cellIs",
						priority: 1,
						operator: "greaterThan",
						formulas: ["1"],
						dxf: { fill: { bgColor: { rgb: "FFFFC7CE" } }, font: { bold: true } },
					},
					{
						type: "colorScale",
						priority: 2,
						cfvo: [{ type: "min" }, { type: "max" }],
						colors: [{ rgb: "FFFFFFFF" }, { rgb: "FF638EC6" }],
					},
					{
						type: "dataBar",
						priority: 3,
						cfvo: [{ type: "min" }, { type: "max" }],
						color: { rgb: "FF638EC6" },
					},
					{
						type: "iconSet",
						priority: 4,
						iconSet: "3TrafficLights1",
						cfvo: [
							{ type: "percent", val: "0" },
							{ type: "percent", val: "33" },
							{ type: "percent", val: "67" },
						],
					},
				],
			},
		];
		expect(await roundTrip(cf)).toEqual(cf);
	});

	it("interns two rules sharing an identical dxf into ONE <dxfs> slot (dedup)", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[1]],
					conditionalFormatting: [
						{
							sqref: ["A1"],
							rules: [
								{
									type: "cellIs",
									priority: 1,
									operator: "equal",
									formulas: ["1"],
									dxf: { font: { bold: true } },
								},
								{
									type: "cellIs",
									priority: 2,
									operator: "equal",
									formulas: ["2"],
									dxf: { font: { bold: true } },
								},
							],
						},
					],
				},
			],
		});
		const styles = await partText(bytes, "xl/styles.xml");
		expect(styles).toContain('<dxfs count="1">'); // one shared dxf, not two
		const sheet = await partText(bytes, "xl/worksheets/sheet1.xml");
		expect((sheet.match(/dxfId="0"/g) ?? []).length).toBe(2); // both rules point at it
	});

	it("does NOT swap a solid dxf fill's colors on round-trip (bgColor stays bgColor)", async () => {
		const [block] = await roundTrip([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "cellIs",
						priority: 1,
						operator: "equal",
						formulas: ["1"],
						dxf: {
							fill: {
								patternType: "solid",
								fgColor: { rgb: "FFAABBCC" },
								bgColor: { rgb: "FFDDEEFF" },
							},
						},
					},
				],
			},
		]);
		const rule = block?.rules[0];
		const fill = rule && "dxf" in rule ? rule.dxf?.fill : undefined;
		expect(fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FFAABBCC" },
			bgColor: { rgb: "FFDDEEFF" },
		});
	});

	it("renumbers priorities densely by ascending caller priority (document order breaks ties)", async () => {
		// Caller priorities are sparse and out of document order; precedence must be PRESERVED.
		const [block] = await roundTrip([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "expression",
						priority: 50,
						formulas: ["TRUE"],
						dxf: { font: { italic: true } },
					},
					{
						type: "expression",
						priority: 10,
						formulas: ["FALSE"],
						dxf: { font: { bold: true } },
					},
					{
						type: "expression",
						priority: 30,
						formulas: ["TRUE"],
						dxf: { font: { strike: true } },
					},
				],
			},
		]);
		const rules = block?.rules ?? [];
		// Renumbered 1..3 by ascending caller priority: 10→1, 30→2, 50→3. The rule that had priority 10
		// (bold) must now be priority 1 (highest precedence).
		const byDxf = (r: (typeof rules)[number]) =>
			r.type === "expression" && r.dxf?.font ? Object.keys(r.dxf.font)[0] : undefined;
		const bold = rules.find((r) => byDxf(r) === "bold");
		const strike = rules.find((r) => byDxf(r) === "strike");
		const italic = rules.find((r) => byDxf(r) === "italic");
		expect(bold?.priority).toBe(1);
		expect(strike?.priority).toBe(2);
		expect(italic?.priority).toBe(3);
	});

	it("emits <conditionalFormatting> between </sheetData> and <dataValidations>", async () => {
		const sheet = await partText(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [[1]],
						conditionalFormatting: [
							{
								sqref: ["A1"],
								rules: [{ type: "expression", priority: 1, formulas: ["TRUE"] }],
							},
						],
						dataValidations: [
							{
								sqref: ["A1"],
								type: "whole",
								operator: "greaterThan",
								formula1: "0",
							},
						],
					},
				],
			}),
			"xl/worksheets/sheet1.xml",
		);
		const cf = sheet.indexOf("<conditionalFormatting");
		const dv = sheet.indexOf("<dataValidations");
		const sd = sheet.indexOf("</sheetData>");
		expect(sd).toBeLessThan(cf);
		expect(cf).toBeLessThan(dv);
	});

	it("emits neither <conditionalFormatting> nor <dxfs> when unused (byte-identity path)", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1]] }] });
		expect(await partText(bytes, "xl/worksheets/sheet1.xml")).not.toContain(
			"conditionalFormatting",
		);
		// A no-style workbook still emits styles.xml? No — needed() is false, so it isn't present.
		expect(openZip(bytes).has("xl/styles.xml")).toBe(false);
	});
});

describe("writeXlsx — conditional formatting rejects malformed input (typed)", () => {
	it("rejects a bad priority, unknown type, empty/non-canonical sqref", async () => {
		await rejects([
			{ sqref: ["A1"], rules: [{ type: "expression", priority: 0, formulas: ["T"] }] },
		]);
		await rejects([{ sqref: ["A1"], rules: [{ type: "bogus", priority: 1 }] }]);
		await rejects([
			{ sqref: [], rules: [{ type: "expression", priority: 1, formulas: ["T"] }] },
		]);
		await rejects([
			{ sqref: ["A:A"], rules: [{ type: "expression", priority: 1, formulas: ["T"] }] },
		]);
		await rejects([{ sqref: ["A1"], rules: [] }]);
	});
	it("rejects out-of-count graphical rules (decision-5 cfvo/color bounds — Excel repair guard)", async () => {
		// colorScale: needs 2-3 cfvo + equal colors.
		await rejects([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "colorScale",
						priority: 1,
						cfvo: [{ type: "min" }],
						colors: [{ rgb: "FFFF0000" }],
					},
				],
			},
		]);
		await rejects([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "colorScale",
						priority: 1,
						cfvo: [{ type: "min" }, { type: "max" }],
						colors: [{ rgb: "FF0000FF" }, { rgb: "FF00FF00" }, { rgb: "FFFF0000" }],
					},
				],
			},
		]);
		// dataBar: needs exactly 2 cfvo.
		await rejects([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "dataBar",
						priority: 1,
						cfvo: [{ type: "min" }],
						color: { rgb: "FF638EC6" },
					},
				],
			},
		]);
		// iconSet: 3TrafficLights1 needs exactly 3 cfvo.
		await rejects([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "iconSet",
						priority: 1,
						iconSet: "3TrafficLights1",
						cfvo: [
							{ type: "percent", val: "0" },
							{ type: "percent", val: "50" },
						],
					},
				],
			},
		]);
	});

	it("rejects an unknown property, a bad dxf, and a malformed color", async () => {
		await rejects([{ sqref: ["A1"], rules: [{ type: "expression", priority: 1, bogus: 1 }] }]);
		await rejects([
			{
				sqref: ["A1"],
				rules: [{ type: "cellIs", priority: 1, dxf: { font: { size: -1 } } }],
			},
		]);
		await rejects([
			{
				sqref: ["A1"],
				rules: [
					{
						type: "colorScale",
						priority: 1,
						cfvo: [{ type: "min" }],
						colors: [{ rgb: "nothex" }],
					},
				],
			},
		]);
	});
});

describe("streamXlsx — conditional formatting lands on the streaming writer too", () => {
	it("emits the same rules through the streamed footer + shared dxfs", async () => {
		const cf: ConditionalFormatting[] = [
			{
				sqref: ["A1:A2"],
				rules: [
					{
						type: "cellIs",
						priority: 1,
						operator: "greaterThan",
						formulas: ["1"],
						dxf: { font: { bold: true } },
					},
				],
			},
		];
		const bytes = await drain(
			streamXlsx({ sheets: [{ name: "S", rows: [[1], [2]], conditionalFormatting: cf }] }),
		);
		const ws = (await openXlsx(bytes)).sheet("S");
		expect(ws.conditionalFormatting).toEqual(cf);
	});
});
