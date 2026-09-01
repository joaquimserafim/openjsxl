import { describe, expect, it } from "vitest";
import { evaluateCell } from "../../formula";
import { MAX_FORMULA_LEN } from "../../ooxml/formula";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F11.20 — ST_Formula is a restriction of ST_Xstring: every supported formula field (cell `<f>`,
// data-validation formula1/2, conditional-formatting <formula>, table totals-row/calculated-column
// formulas) is LOGICAL text on both sides. The writer encodes the complete value exactly once,
// after validation and immediately before XML escaping; the reader decodes once. A literal
// `_x0041_` therefore stores protected as `_x005F_x0041_` (Excel would read a raw `_x0041_` as `A`)
// and round-trips literally; clean formulas are byte-identical; CR travels as `_x000D_` (E20).

const decoder = new TextDecoder();
const part = async (bytes: Uint8Array, path: string): Promise<string> =>
	decoder.decode(await openZip(bytes).read(path));
const sheet1 = (bytes: Uint8Array) => part(bytes, "xl/worksheets/sheet1.xml");

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
	let o = 0;
	for (const c of chunks) {
		out.set(c, o);
		o += c.length;
	}
	return out;
}

// A sheet reference AND a string constant, both escape-shaped — the two places Excel's decoding
// silently changes meaning (a dangling sheet ref; `"A"` instead of `"_x0041_"`).
const LOGICAL = `'_x0041_'!A1&"_x0041_"`;
const WIRE = `'_x005F_x0041_'!A1&amp;"_x005F_x0041_"`;

describe("writeXlsx — ST_Formula fields are ST_Xstring (F11.20)", () => {
	it("cell <f>: stores a logical escape-shaped formula protected and reads it back literally — both writers", async () => {
		const input: WorkbookInput = {
			sheets: [{ name: "S", rows: [[{ formula: LOGICAL, value: 1 }]] }],
		};
		const buffered = await writeXlsx(input);
		expect(await sheet1(buffered)).toContain(`<f>${WIRE}</f>`);
		expect((await openXlsx(buffered)).sheet("S").formula("A1")).toBe(LOGICAL);

		const streamed = await drain(streamXlsx(input));
		expect(await sheet1(streamed)).toContain(`<f>${WIRE}</f>`); // identical affected XML (D5)
		expect((await openXlsx(streamed)).sheet("S").formula("A1")).toBe(LOGICAL);
	});

	it("data validation, conditional formatting, and table formulas: same encode-once/decode-once contract", async () => {
		const input: WorkbookInput = {
			sheets: [
				{
					name: "S",
					rows: [
						["Item", "Amount"],
						["a", 1],
						["Total", 1],
					],
					dataValidations: [
						{ sqref: ["C1"], type: "custom", formula1: LOGICAL, formula2: "_x0041_" },
					],
					conditionalFormatting: [
						{
							sqref: ["C1"],
							rules: [{ type: "expression", priority: 1, formulas: [LOGICAL] }],
						},
					],
					tables: [
						{
							name: "T",
							ref: "A1:B3",
							columns: [
								{ name: "Item" },
								{
									name: "Amount",
									totalsRowFormula: "SUBTOTAL(109,[_x0041_])",
									calculatedColumnFormula: '[Item]&"_x0041_"',
								},
							],
							headerRow: true,
							totalsRow: true,
						},
					],
				},
			],
		};
		const bytes = await writeXlsx(input);
		const xml = await sheet1(bytes);
		expect(xml).toContain(`<formula1>${WIRE}</formula1><formula2>_x005F_x0041_</formula2>`);
		expect(xml).toContain(`<formula>${WIRE}</formula>`);
		const table = await part(bytes, "xl/tables/table1.xml");
		expect(table).toContain(
			"<totalsRowFormula>SUBTOTAL(109,[_x005F_x0041_])</totalsRowFormula>",
		);
		expect(table).toContain(
			'<calculatedColumnFormula>[Item]&amp;"_x005F_x0041_"</calculatedColumnFormula>',
		);

		const s = (await openXlsx(bytes)).sheet("S");
		expect(s.dataValidations[0]?.formula1).toBe(LOGICAL);
		expect(s.dataValidations[0]?.formula2).toBe("_x0041_");
		const rule = s.conditionalFormatting[0]?.rules[0];
		expect(rule !== undefined && "formulas" in rule ? rule.formulas : undefined).toEqual([
			LOGICAL,
		]);
		expect(s.tables[0]?.columns[1]?.totalsRowFormula).toBe("SUBTOTAL(109,[_x0041_])");
		expect(s.tables[0]?.columns[1]?.calculatedColumnFormula).toBe('[Item]&"_x0041_"');
	});

	it("DV/CF/table: the streaming writer emits the same affected parts, and the bridge re-emits them", async () => {
		const input: WorkbookInput = {
			sheets: [
				{
					name: "S",
					rows: [
						["Item", "Amount"],
						["a", 1],
						["Total", 1],
					],
					dataValidations: [{ sqref: ["C1"], type: "custom", formula1: LOGICAL }],
					conditionalFormatting: [
						{
							sqref: ["C1"],
							rules: [{ type: "expression", priority: 1, formulas: [LOGICAL] }],
						},
					],
					tables: [
						{
							name: "T",
							ref: "A1:B3",
							columns: [
								{ name: "Item" },
								{ name: "Amount", totalsRowFormula: "SUBTOTAL(109,[_x0041_])" },
							],
							headerRow: true,
							totalsRow: true,
						},
					],
				},
			],
		};
		const buffered = await writeXlsx(input);
		const streamed = await drain(streamXlsx(input));
		// The streamed sheet part omits <dimension> by design; compare the affected blocks only.
		const block = (xml: string, open: string, close: string) =>
			xml.slice(xml.indexOf(open), xml.indexOf(close) + close.length);
		const b = await sheet1(buffered);
		const s = await sheet1(streamed);
		expect(block(s, "<dataValidations", "</dataValidations>")).toBe(
			block(b, "<dataValidations", "</dataValidations>"),
		);
		expect(block(s, "<conditionalFormatting", "</conditionalFormatting>")).toBe(
			block(b, "<conditionalFormatting", "</conditionalFormatting>"),
		);
		expect(await part(streamed, "xl/tables/table1.xml")).toBe(
			await part(buffered, "xl/tables/table1.xml"),
		);

		const bridged = await writeXlsx(await workbookToInput(await openXlsx(buffered)));
		expect(await sheet1(bridged)).toContain(`<formula1>${WIRE}</formula1>`);
		expect(await sheet1(bridged)).toContain(`<formula>${WIRE}</formula>`);
		expect(await part(bridged, "xl/tables/table1.xml")).toContain("[_x005F_x0041_]");
	});

	it("keeps clean formula bytes identical — encoding is the identity when nothing is escaped", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ formula: "SUM(A1:A9)*2", value: 0 },
							{ formula: "A1<B1", value: true },
						],
					],
				},
			],
		});
		const xml = await sheet1(bytes);
		expect(xml).toContain("<f>SUM(A1:A9)*2</f>");
		expect(xml).toContain("<f>A1&lt;B1</f>");
		expect(xml).not.toContain("_x00"); // no escape introduced anywhere
	});

	it("validates the LOGICAL length: a formula within the ceiling writes even when its wire form exceeds it", async () => {
		// 1170 literal `_x0041_` segments: logical 8,190 chars (≤ MAX_FORMULA_LEN), wire 15,210 chars.
		const logical = "_x0041_".repeat(1170);
		expect(logical.length).toBeLessThanOrEqual(MAX_FORMULA_LEN);
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[{ formula: logical, value: 0 }]] }],
		});
		expect((await sheet1(bytes)).length).toBeGreaterThan(MAX_FORMULA_LEN);
		expect((await openXlsx(bytes)).sheet("S").formula("A1")).toBe(logical);
		// …while a logical formula OVER the ceiling is still refused, regardless of wire size.
		await expect(
			writeXlsx({
				sheets: [
					{ name: "S", rows: [[{ formula: "A".repeat(MAX_FORMULA_LEN + 1), value: 0 }]] },
				],
			}),
		).rejects.toMatchObject({ code: "invalid-input" });
	});

	it("bridges read → workbookToInput → write → read losslessly for escape-shaped formulas", async () => {
		const first = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[{ formula: LOGICAL, value: 1 }]],
					dataValidations: [{ sqref: ["B1"], type: "custom", formula1: "_x0041_" }],
				},
			],
		});
		const second = await writeXlsx(await workbookToInput(await openXlsx(first)));
		const s = (await openXlsx(second)).sheet("S");
		expect(s.formula("A1")).toBe(LOGICAL);
		expect(s.dataValidations[0]?.formula1).toBe("_x0041_");
		// The bridge re-emits the protected wire form, never a raw `_x0041_`.
		expect(await sheet1(second)).toContain(`<f>${WIRE}</f>`);
	});

	it("the evaluator resolves an escape-shaped sheet reference to the escape-shaped sheet", async () => {
		// Tab `_x0041_` (F11.4) and the formula's `'_x0041_'!A1` (F11.20) are encoded by the same
		// codec, so after decoding they name the SAME sheet — the coupling the E20 review proved.
		const bytes = await writeXlsx({
			sheets: [
				{ name: "_x0041_", rows: [[21]] },
				{ name: "S", rows: [[{ formula: "'_x0041_'!A1*2", value: 0 }]] },
			],
		});
		expect(await evaluateCell(await openXlsx(bytes), "S", "A1")).toBe(42);
	});
});
