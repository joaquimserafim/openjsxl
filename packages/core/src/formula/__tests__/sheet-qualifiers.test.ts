import { describe, expect, it, onTestFinished, vi } from "vitest";
import type { DefinedName } from "../../ooxml/workbook";
import { openOds } from "../../reader/ods";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import { writeZip } from "../../writer/zip";
import { evaluateCell, evaluateWorkbook } from "../eval";
import { errorValue } from "../value";
import { evaluateFormulas } from "./criteria-fixtures";

// F11.65: Data (sheet 0) holds 5 and 7 in A1:A2 and 4 in B1; the formulas live on Results (sheet 1).
const data = [
	[5, 4],
	[7, null],
];

// Rate is scoped to Data and global; Tax is scoped to Data only; Fee to Results only; Global and
// Ref are workbook-wide.
const names: readonly DefinedName[] = [
	{ name: "Rate", refersTo: "0.07", localSheetId: 0 },
	{ name: "Rate", refersTo: "0.5" },
	{ name: "Tax", refersTo: "0.2", localSheetId: 0 },
	{ name: "Fee", refersTo: "3", localSheetId: 1 },
	{ name: "Global", refersTo: "9" },
	{ name: "Ref", refersTo: "data!$A$2" },
];

describe("a sheet-qualified defined name", () => {
	it("uses the name scoped to that sheet before the workbook's", async () => {
		const [scoped, global, percent] = await evaluateFormulas(
			["Data!Rate", "Rate", "Data!Rate*100"],
			data,
			names,
		);
		expect([scoped, global]).toEqual([0.07, 0.5]);
		expect(percent).toBeCloseTo(7, 12);
	});

	it("finds a name scoped only to that sheet from another sheet", async () => {
		expect(await evaluateFormulas(["Data!Tax", "Tax"], data, names)).toEqual([
			0.2,
			errorValue("#NAME?"),
		]);
	});

	it("falls back to the workbook's name, but never to another sheet's", async () => {
		expect(await evaluateFormulas(["Data!Global", "Data!Fee", "Fee"], data, names)).toEqual([
			9,
			errorValue("#NAME?"),
			3,
		]);
	});

	it("gives #REF! when the qualifier names no sheet, whether or not the name exists", async () => {
		expect(
			await evaluateFormulas(["Nope!Rate", "Nope!Missing", "Data!Missing"], data, names),
		).toEqual([errorValue("#REF!"), errorValue("#REF!"), errorValue("#NAME?")]);
	});

	it("matches its sheet ignoring case", async () => {
		expect(await evaluateFormulas(["data!Rate", "'DATA'!Tax"], data, names)).toEqual([
			0.07, 0.2,
		]);
	});
});

describe("a sheet qualifier on a reference", () => {
	it("matches its sheet ignoring case, quoted or not", async () => {
		expect(
			await evaluateFormulas(["data!A1", "DATA!A1+1", "'data'!A1", "dAtA!B1"], data),
		).toEqual([5, 6, 5, 4]);
	});

	it("does the same on a range, a whole column and a whole row", async () => {
		expect(
			await evaluateFormulas(["SUM(data!A1:A2)", "SUM(DATA!A:A)", "SUM(data!1:1)"], data),
		).toEqual([12, 12, 9]);
	});

	it("reaches a defined name's reference whose sheet is spelled in another case", async () => {
		expect(await evaluateFormulas(["Ref", "Ref*2"], data, names)).toEqual([7, 14]);
	});

	it("still gives #REF! for a sheet that does not exist", async () => {
		expect(await evaluateFormulas(["Nope!A1", "SUM(Nope!A1:A2)"], data)).toEqual([
			errorValue("#REF!"),
			errorValue("#REF!"),
		]);
	});

	it("gives #REF! when either end of a range names no sheet", async () => {
		const formulas = [
			"SUM(Data!A1:Nope!A2)",
			"SUM(Nope!A1:Data!A2)",
			"SUM(Data!A:Nope!A)",
			"SUM(Data!1:Nope!2)",
			"SUM(A1:Nope!A2)",
		];
		expect(await evaluateFormulas(formulas, data)).toEqual(
			formulas.map(() => errorValue("#REF!")),
		);
	});

	it("accepts a qualifier on both ends of a range, each in any case", async () => {
		expect(
			await evaluateFormulas(
				["SUM(Data!A1:data!A2)", "SUM(data!A:DATA!A)", "SUM(DATA!1:data!1)"],
				data,
			),
		).toEqual([12, 12, 9]);
	});

	it("is the same cell under every spelling, so a cycle through it is still found", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "Data",
						rows: [
							[
								{ formula: "data!B1+1", value: 0 },
								{ formula: "DATA!A1+1", value: 0 },
							],
						],
					},
				],
			}),
		);
		expect(await evaluateCell(book, "Data", "A1")).toEqual(errorValue("#CYCLE!"));
	});
});

describe("evaluateCell's sheet name", () => {
	async function book() {
		return openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Data", rows: [[5]] },
					{ name: "Main", rows: [[{ formula: "Data!A1*2", value: 0 }]] },
				],
			}),
		);
	}

	it("matches ignoring case", async () => {
		const wb = await book();
		expect(await evaluateCell(wb, "main", "A1")).toBe(10);
		expect(await evaluateCell(wb, "MAIN", "A1")).toBe(10);
		expect(await evaluateCell(wb, "data", "A1")).toBe(5);
	});

	it("gives #REF! for a sheet that does not exist", async () => {
		expect(await evaluateCell(await book(), "Nope", "A1")).toEqual(errorValue("#REF!"));
	});

	it("gives #REF! for a sheet name that is not a string, as before", async () => {
		const wb = await book();
		for (const sheet of [undefined, null, 123, {}]) {
			expect(await evaluateCell(wb, sheet as never, "A1")).toEqual(errorValue("#REF!"));
		}
	});
});

describe("the sheets a qualifier can find", () => {
	it("are the sheets that loaded, whatever another read of the sheet list says", async () => {
		const real = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "Data", rows: [[5, { formula: "Ghost!A1", value: 0 }]] }],
			}),
		);
		// The first read of `sheets` also lists Ghost, which has no worksheet; later reads do not.
		let reads = 0;
		const ghost = { name: "Ghost", path: "", visible: true, state: "visible" };
		const book = new Proxy(real, {
			get(target, key) {
				if (key === "sheets")
					return ++reads === 1 ? [...target.sheets, ghost] : target.sheets;
				const value = Reflect.get(target, key, target);
				return typeof value === "function" ? value.bind(target) : value;
			},
		});
		const result = await evaluateWorkbook(book);
		expect(result.get("Data", "B1")).toEqual(errorValue("#REF!"));
	});

	it("are found without reading the sheet list again", async () => {
		// Count reads of `sheets` while evaluating one formula; a qualifier must add none.
		async function sheetReads(formula: string) {
			const real = await openXlsx(
				await writeXlsx({ sheets: [{ name: "Data", rows: [[5, { formula, value: 0 }]] }] }),
			);
			let reads = 0;
			const book = new Proxy(real, {
				get(target, key) {
					if (key === "sheets") reads++;
					const value = Reflect.get(target, key, target);
					return typeof value === "function" ? value.bind(target) : value;
				},
			});
			const result = await evaluateWorkbook(book);
			return { reads, value: result.get("Data", "B1") };
		}
		const plain = await sheetReads("A1");
		expect(plain.value).toBe(5);
		for (const formula of ["Data!A1", "data!A1", "SUM(DATA!A1:A1)", "Nope!A1+0*A1"]) {
			const qualified = await sheetReads(formula);
			expect(qualified.reads).toBe(plain.reads);
		}
	});

	it("are matched by folding each long qualifier once, however many references use it", async () => {
		const long = "x".repeat(8000);
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "Data",
						rows: Array.from({ length: 20 }, () => [
							{
								formula: `SUM(${Array.from({ length: 50 }, () => "N").join(",")})`,
								value: 0,
							},
						]),
					},
				],
				definedNames: [{ name: "N", refersTo: `'${long}'!$A$1` }],
			}),
		);
		const lower = vi.spyOn(String.prototype, "toLowerCase");
		onTestFinished(() => lower.mockRestore());
		const result = await evaluateWorkbook(book);
		expect(result.cells.map((cell) => cell.value)).toEqual(
			Array.from({ length: 20 }, () => errorValue("#REF!")),
		);
		const folds = lower.mock.contexts.filter((context) => String(context) === long);
		expect(folds.length).toBeLessThanOrEqual(1);
	});
});

describe("two sheets whose names differ only in case", () => {
	// The xlsx reader keeps the first of such a pair; the ods reader keeps both.
	async function odsBook() {
		const table = (name: string, value: number) =>
			`<table:table table:name="${name}"><table:table-row><table:table-cell office:value-type="float" office:value="${value}"/></table:table-row></table:table>`;
		const content =
			'<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
			'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3">' +
			`<office:body><office:spreadsheet>${table("Data", 1)}${table("data", 2)}</office:spreadsheet></office:body></office:document-content>`;
		const encoder = new TextEncoder();
		return openOds(
			await writeZip([
				{
					name: "mimetype",
					data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"),
				},
				{ name: "content.xml", data: encoder.encode(content) },
			]),
		);
	}

	it("keeps each reachable by its exact spelling, and any other spelling finds the first", async () => {
		const wb = await odsBook();
		expect(wb.sheets.map((s) => s.name)).toEqual(["Data", "data"]);
		expect(await evaluateCell(wb, "Data", "A1")).toBe(1);
		expect(await evaluateCell(wb, "data", "A1")).toBe(2);
		expect(await evaluateCell(wb, "DATA", "A1")).toBe(1);
	});
});

describe("what F11.65 leaves as it was", () => {
	it("a range whose ends name two different sheets reads the left one (recorded)", async () => {
		expect(
			await evaluateFormulas(["SUM(Data!A1:Results!A2)", "SUM(A1:Data!A2)"], data),
		).toEqual([12, 12]);
	});

	it("evaluateWorkbook's get matches the sheet as cells spells it", async () => {
		const book = await openXlsx(
			await writeXlsx({ sheets: [{ name: "Main", rows: [[{ formula: "1+1", value: 0 }]] }] }),
		);
		const result = await evaluateWorkbook(book);
		expect(result.cells).toEqual([{ sheet: "Main", ref: "A1", value: 2 }]);
		expect([result.get("Main", "A1"), result.get("main", "A1")]).toEqual([2, undefined]);
	});

	it("an unqualified name resolves on its own sheet, then the workbook's", async () => {
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Data", rows: [[{ formula: "Rate", value: 0 }]] },
					{ name: "Results", rows: [[{ formula: "Rate", value: 0 }]] },
				],
				definedNames: names,
			}),
		);
		expect(await evaluateCell(book, "Data", "A1")).toBe(0.07);
		expect(await evaluateCell(book, "Results", "A1")).toBe(0.5);
	});

	it("a 3-D name or reference is #REF!", async () => {
		expect(
			await evaluateFormulas(["Data:Results!Rate", "SUM(Data:Results!A1)"], data, names),
		).toEqual([errorValue("#REF!"), errorValue("#REF!")]);
	});
});
