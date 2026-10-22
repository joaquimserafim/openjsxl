import { createRequire } from "node:module";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { openZip } from "../../../core/src/zip";
import * as sheetJsAdapter from "../adapters/xlsx.mjs";

const UUID_V4 = /^\{[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\}$/i;

function ruleIds(xml: string) {
	return Array.from(xml.matchAll(/<x14:cfRule\b[^>]*\bid="([^"]+)"/g), (match) => match[1]);
}

function expectUniqueIds(xml: string, count: number) {
	const ids = ruleIds(xml);
	expect(ids).toHaveLength(count);
	expect(new Set(ids).size).toBe(count);
	for (const id of ids) expect(id).toMatch(UUID_V4);
}

// Normal Workbook writes prepare every rule, so the renderer's second uuid.v4() call only runs
// when rendering an unprepared icon rule. Exercise that compatibility surface directly as well.
function unpreparedIconXml(): string {
	const ruleTransform: unknown = createRequire(import.meta.url)(
		"exceljs/lib/xlsx/xform/sheet/cf-ext/cf-rule-ext-xform.js",
	);
	if (typeof ruleTransform !== "function")
		throw new Error("ExcelJS rule transform is not available");
	const transform: unknown = Reflect.construct(ruleTransform, []);
	if (
		typeof transform !== "object" ||
		transform === null ||
		!("toXml" in transform) ||
		typeof transform.toXml !== "function"
	) {
		throw new Error("ExcelJS rule transform cannot render XML");
	}
	const xml: unknown = transform.toXml({
		type: "iconSet",
		iconSet: "3Stars",
		priority: 1,
		cfvo: [
			{ type: "percent", value: 0 },
			{ type: "percent", value: 33 },
			{ type: "percent", value: 67 },
		],
	});
	if (typeof xml !== "string") throw new Error("ExcelJS rule transform returned non-text XML");
	return xml;
}

describe("reviewed benchmark comparator dependencies", () => {
	it("round-trips ExcelJS x14 data bars and icon sets with distinct valid UUIDs", async () => {
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet("UUIDs");
		worksheet.addRows([
			[1, 10],
			[2, 20],
			[3, 30],
		]);
		worksheet.addConditionalFormatting({
			ref: "A1:A3",
			rules: [
				{
					type: "dataBar",
					priority: 1,
					gradient: false,
					cfvo: [{ type: "min" }, { type: "max" }],
				},
			],
		});
		worksheet.addConditionalFormatting({
			ref: "B1:B3",
			rules: [
				{
					type: "iconSet",
					priority: 2,
					iconSet: "3Stars",
					cfvo: [
						{ type: "percent", value: 0 },
						{ type: "percent", value: 33 },
						{ type: "percent", value: 67 },
					],
				},
			],
		});

		const bytes = await workbook.xlsx.writeBuffer();
		const xml = new TextDecoder().decode(
			await openZip(new Uint8Array(bytes)).read("xl/worksheets/sheet1.xml"),
		);
		expectUniqueIds(xml, 2);
		expect(xml).toContain("<x14:dataBar");
		expect(xml).toContain('iconSet="3Stars"');

		const reopened = new ExcelJS.Workbook();
		await reopened.xlsx.load(bytes);
		const restored = reopened.getWorksheet("UUIDs");
		expect(restored?.getCell("A3").value).toBe(3);
		expect(restored?.getCell("B3").value).toBe(30);
		expect(restored).toMatchObject({
			conditionalFormattings: expect.arrayContaining([
				expect.objectContaining({
					ref: "A1:A3",
					rules: expect.arrayContaining([
						expect.objectContaining({ type: "dataBar", gradient: false }),
					]),
				}),
				expect.objectContaining({
					ref: "B1:B3",
					rules: expect.arrayContaining([
						expect.objectContaining({ type: "iconSet", iconSet: "3Stars" }),
					]),
				}),
			]),
		});
	});

	it("retains the ExcelJS icon renderer's unprepared UUID fallback", () => {
		expectUniqueIds(unpreparedIconXml() + unpreparedIconXml(), 2);
	});

	it("reads an independent buffer fixture through the SheetJS adapter, including cached formulas", async () => {
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet("Mixed");
		worksheet.addRow([12.5, "plain", true, false, new Date("2020-01-02T00:00:00.000Z")]);
		worksheet.addRow([
			{ formula: "SUM(2,3)", result: 5 },
			{ formula: '"cached"', result: "cached" },
			{ formula: "1=1", result: true },
		]);
		const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());

		// The buffer reader uses raw values: 2020-01-02 is serial 43832; cached results contribute
		// 5, "cached".length, and true. A lost cached result or date conversion changes this sink.
		expect(await sheetJsAdapter.read(bytes)).toBe(43862.5);
	});

	it("writes numbers, strings, booleans, and UTC dates through the SheetJS buffer adapter", async () => {
		const date = new Date("2020-01-02T00:00:00.000Z");
		const bytes = await sheetJsAdapter.write([[12.5, "A & Ω", true, false, date]], "numbers");
		expect(bytes).toBeInstanceOf(Uint8Array);

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(new Uint8Array(bytes).buffer);
		const worksheet = workbook.getWorksheet("Bench");
		expect(worksheet?.getCell("A1").value).toBe(12.5);
		expect(worksheet?.getCell("B1").value).toBe("A & Ω");
		expect(worksheet?.getCell("C1").value).toBe(true);
		expect(worksheet?.getCell("D1").value).toBe(false);
		expect(worksheet?.getCell("E1").value).toEqual(date);
		expect(await sheetJsAdapter.read(bytes)).toBe(43850.5);
	});
});
