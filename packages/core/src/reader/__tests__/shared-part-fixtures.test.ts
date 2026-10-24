import { describe, expect, it } from "vitest";
import { openZip } from "../../zip";
import { openXlsx } from "../workbook";
import {
	declaredSheets,
	relationships,
	rewriteParts,
	TABLE_REL,
	threeSheetWorkbook,
} from "./shared-part-fixtures";

describe("shared-part fixtures", () => {
	it("builds the three-sheet workbook", async () => {
		const book = await openXlsx(await threeSheetWorkbook());
		expect(book.sheets.map((s) => s.name)).toEqual(["S1", "S2", "S3"]);
		expect(book.definedNames.map((n) => n.name)).toEqual(["Everywhere", "OnFirst", "OnThird"]);
		expect([book.sheet("S1").tables.length, book.sheet("S2").comments.length]).toEqual([1, 1]);
	});

	it("rewrites a part's text and keeps every other part", async () => {
		const base = await threeSheetWorkbook();
		const edited = await rewriteParts(base, {
			"xl/workbook.xml": (t) => t.replace('name="S3"', 'name="Third"'),
		});
		expect((await openXlsx(edited)).sheets.map((s) => s.name)).toEqual(["S1", "S2", "Third"]);
		expect([...openZip(edited).entries.keys()]).toEqual([...openZip(base).entries.keys()]);
	});

	it("refuses a missing part and an edit that changes nothing", async () => {
		const base = await threeSheetWorkbook();
		await expect(rewriteParts(base, { "missing.xml": (t) => `${t}x` })).rejects.toThrow(
			"missing.xml is not in the archive",
		);
		await expect(rewriteParts(base, { "xl/workbook.xml": (t) => t })).rejects.toThrow(
			"the edit left xl/workbook.xml unchanged",
		);
	});

	it("names the OPC table relationship type", () => {
		expect(TABLE_REL.endsWith("/relationships/table")).toBe(true);
	});

	it("reads the declared sheets and the parts their relationships name from raw XML", async () => {
		expect(await declaredSheets(await threeSheetWorkbook())).toEqual([
			{
				path: "xl/worksheets/sheet1.xml",
				tables: ["xl/tables/table1.xml"],
				comments: "xl/comments1.xml",
				commentCount: 1,
			},
			{
				path: "xl/worksheets/sheet2.xml",
				tables: [],
				comments: "xl/comments2.xml",
				commentCount: 1,
			},
			{ path: "xl/worksheets/sheet3.xml", tables: [], comments: undefined, commentCount: 0 },
		]);
		await expect(declaredSheets(new Uint8Array(4))).rejects.toMatchObject({
			code: "not-a-zip",
		});
	});

	it("parses relationship elements", () => {
		expect(
			relationships(
				'<Relationship Id="rId1" Type="t/worksheet" Target="a.xml"/><Relationship Id="rId2" Type="t/hyperlink" Target="https://x/" TargetMode="External"/>',
			),
		).toEqual([
			{ id: "rId1", type: "t/worksheet", target: "a.xml", external: false },
			{ id: "rId2", type: "t/hyperlink", target: "https://x/", external: true },
		]);
		expect(relationships("<Relationships/>")).toEqual([]);
	});
});
