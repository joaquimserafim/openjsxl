import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import type { SheetInput, WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

const ordinaryInputs: readonly { label: string; workbook: WorkbookInput }[] = [
	{
		label: "no dates",
		workbook: {
			sheets: [
				{
					name: "S",
					rows: [
						["text", 7, true],
						[null, 0, false],
					],
				},
			],
		},
	},
	{
		label: "1900 dates",
		workbook: { sheets: [{ name: "S", rows: [[new Date("2024-02-29T12:00:00Z")]] }] },
	},
	{
		label: "multiple sparse sheets",
		workbook: {
			sheets: [
				{ name: "First", rows: [undefined, [null, 7], undefined, ["last"]] },
				{ name: "Second", rows: [[true, undefined, "second"]], state: "hidden" },
			],
		},
	},
	{
		label: "1904 dates",
		workbook: {
			date1904: true,
			sheets: [{ name: "S", rows: [[new Date("2024-02-29T12:00:00Z")]] }],
		},
	},
	{ label: "an empty sheet", workbook: { sheets: [{ name: "S", rows: [] }] } },
];

describe("writeXlsx — sheet.rows snapshot", () => {
	it("emits the first rows array when a getter later returns a different array", async () => {
		let reads = 0;
		const sheet: SheetInput = {
			name: "S",
			get rows() {
				reads++;
				return reads === 1 ? [[7]] : [[99]];
			},
		};

		const workbook = await openXlsx(await writeXlsx({ sheets: [sheet] }));

		expect(workbook.sheet("S").cell("A1").value).toBe(7);
		expect(reads).toBe(1);
	});

	it("does not read a later null value after validating an array", async () => {
		let reads = 0;
		const sheet: SheetInput = { name: "S", rows: [] };
		Object.defineProperty(sheet, "rows", {
			get() {
				reads++;
				return reads === 1 ? [[7]] : null;
			},
		});

		const workbook = await openXlsx(await writeXlsx({ sheets: [sheet] }));

		expect(workbook.sheet("S").cell("A1").value).toBe(7);
		expect(reads).toBe(1);
	});

	it("does not invoke a getter that would throw on its second access", async () => {
		let reads = 0;
		const sheet: SheetInput = {
			name: "S",
			get rows() {
				reads++;
				if (reads > 1) throw new Error("rows was read again");
				return [[7]];
			},
		};

		const workbook = await openXlsx(await writeXlsx({ sheets: [sheet] }));

		expect(workbook.sheet("S").cell("A1").value).toBe(7);
		expect(reads).toBe(1);
	});

	it("retains an empty first array instead of emitting later rows", async () => {
		let reads = 0;
		const sheet: SheetInput = {
			name: "S",
			get rows() {
				reads++;
				return reads === 1 ? [] : [["unexpected"]];
			},
		};

		const workbook = await openXlsx(await writeXlsx({ sheets: [sheet] }));

		expect(workbook.sheet("S").cell("A1").value).toBeNull();
		expect(reads).toBe(1);
	});

	it("gets rows once on each sheet Proxy, including empty and sparse sheets", async () => {
		const reads: string[] = [];
		const sheets: SheetInput[] = [
			{ name: "First", rows: [[7]] },
			{ name: "Empty", rows: [] },
			{ name: "Third", rows: [undefined, ["third"]] },
		];
		const proxies = sheets.map(
			(sheet) =>
				new Proxy(sheet, {
					get(target, key, receiver) {
						if (key === "rows") {
							if (reads.includes(target.name)) throw new Error("repeated rows get");
							reads.push(target.name);
						}
						return Reflect.get(target, key, receiver);
					},
				}),
		);

		const workbook = await openXlsx(await writeXlsx({ sheets: proxies }));

		expect(reads).toEqual(["First", "Empty", "Third"]);
		expect(workbook.sheet("First").cell("A1").value).toBe(7);
		expect(workbook.sheet("Empty").cell("A1").value).toBeNull();
		expect(workbook.sheet("Third").cell("A2").value).toBe("third");
	});

	it.each([
		{ label: "null", value: null },
		{ label: "undefined", value: undefined },
		{ label: "a string", value: "rows" },
		{ label: "a non-array object", value: {} },
		{ label: "a sync iterable", value: new Set([[7]]) },
	])("rejects $label on the first read with a typed error", async ({ value }) => {
		let reads = 0;
		const sheet: SheetInput = { name: "S", rows: [] };
		Object.defineProperty(sheet, "rows", {
			get() {
				reads++;
				return reads === 1 ? value : [[7]];
			},
		});

		const result = writeXlsx({ sheets: [sheet] });

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({
			code: "invalid-input",
			message: /rows must be an array/,
		});
		expect(reads).toBe(1);
	});

	it("validates the captured array contents even if a later array would be valid", async () => {
		let reads = 0;
		const sheet: SheetInput = {
			name: "S",
			get rows() {
				reads++;
				return reads === 1 ? [[Number.POSITIVE_INFINITY]] : [[7]];
			},
		};

		const result = writeXlsx({ sheets: [sheet] });

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({ code: "invalid-input" });
		expect(reads).toBe(1);
	});

	it("derives table headers and worksheet bounds from the captured rows", async () => {
		let reads = 0;
		const sheet: SheetInput = {
			name: "S",
			get rows() {
				reads++;
				return reads === 1
					? [
							["Region", "Total"],
							["West", 10],
							["East", 20],
						]
					: [["Wrong", "Header", "Extra"], ["replacement"], [], ["extra row"]];
			},
			tables: [
				{ name: "Sales", ref: "A1:B3", columns: [], headerRow: true, totalsRow: false },
			],
		};

		const workbook = await openXlsx(await writeXlsx({ sheets: [sheet] }));
		const worksheet = workbook.sheet("S");

		expect(worksheet.tables[0]?.columns).toEqual([{ name: "Region" }, { name: "Total" }]);
		expect(worksheet.cell("A1").value).toBe("Region");
		expect(worksheet.cell("B1").value).toBe("Total");
		expect(worksheet.cell("B3").value).toBe(20);
		expect(worksheet.dimension).toBe("A1:B3");
		expect(reads).toBe(1);
	});

	it.each(ordinaryInputs)("preserves plain-input bytes for $label", async ({ workbook }) => {
		const plain = await writeXlsx(workbook);
		const reads: number[] = [];
		const sheets = workbook.sheets.map((sheet, index) => {
			const { rows, ...metadata } = sheet;
			return {
				...metadata,
				get rows() {
					reads[index] = (reads[index] ?? 0) + 1;
					return rows;
				},
			};
		});

		const captured = await writeXlsx({ ...workbook, sheets });

		expect(captured).toEqual(plain);
		expect(reads).toEqual(sheets.map(() => 1));
	});
});
