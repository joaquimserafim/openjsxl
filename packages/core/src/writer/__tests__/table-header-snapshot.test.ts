import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import type { CellStyle, TableInfo } from "../../types";
import { streamXlsx } from "../stream";
import type { CellInput, CellValue, SheetInput, StyledCell } from "../types";
import { writeXlsx } from "../workbook";

function table(name = "Data", ref = "A1:A2"): TableInfo {
	return { name, ref, columns: [], headerRow: true, totalsRow: false };
}

async function roundTrip(rows: SheetInput["rows"], tables = [table()]) {
	return (await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows, tables }] }))).sheet("S");
}

describe("writeXlsx — table header value snapshot", () => {
	it("shares the first getter value between inferred name and emitted header", async () => {
		let reads = 0;
		const header: CellInput = {
			get value() {
				reads++;
				return reads === 1 ? "Header" : "Different";
			},
		};

		const sheet = await roundTrip([[header], [42]]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(reads).toBe(1);
	});

	it("does not reread a header value that throws after the first access", async () => {
		let reads = 0;
		const header: StyledCell = {
			get value() {
				if (++reads > 1) throw new Error("header value read twice");
				return "Header";
			},
		};

		const sheet = await roundTrip([[header], [42]]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(reads).toBe(1);
	});

	it("performs one Proxy value get and retains the first string", async () => {
		let reads = 0;
		const header = new Proxy(
			{ value: "Header" },
			{
				get(target, key, receiver) {
					if (key === "value") return ++reads === 1 ? "Header" : "Different";
					return Reflect.get(target, key, receiver);
				},
			},
		);

		const sheet = await roundTrip([[header], [42]]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(reads).toBe(1);
	});

	it("preserves a cached formula string header and captures formula text once", async () => {
		let valueReads = 0;
		let formulaReads = 0;
		const header: StyledCell = {
			get value() {
				if (++valueReads > 1) throw new Error("cached value read twice");
				return "Header";
			},
			get formula() {
				if (++formulaReads > 1) throw new Error("formula read twice");
				return '"Header"';
			},
		};

		const sheet = await roundTrip([[header], [42]]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(sheet.formula("A1")).toBe('"Header"');
		expect(valueReads).toBe(1);
		expect(formulaReads).toBe(1);
	});

	it("captures distinct header coordinates independently across several tables", async () => {
		const reads = [0, 0, 0];
		function header(index: number, value: string): StyledCell {
			return {
				get value() {
					reads[index] = (reads[index] ?? 0) + 1;
					return reads[index] === 1 ? value : "Changed";
				},
			};
		}
		const sheet = await roundTrip(
			[
				[header(0, "Left"), undefined, header(1, "Right")],
				[1, undefined, 2],
				undefined,
				[header(2, "Lower")],
				[3],
			],
			[table("LeftData"), table("RightData", "C1:C2"), table("LowerData", "A4:A5")],
		);

		expect(sheet.tables.map((t) => t.columns[0]?.name)).toEqual(["Left", "Right", "Lower"]);
		expect(["A1", "C1", "A4"].map((ref) => sheet.cell(ref).value)).toEqual([
			"Left",
			"Right",
			"Lower",
		]);
		expect(reads).toEqual([1, 1, 1]);
	});

	it("rejects overlapping tables after sharing their common header capture", async () => {
		let reads = 0;
		const shared: StyledCell = {
			get value() {
				if (++reads > 1) throw new Error("overlapping consumer reread the header");
				return "Shared";
			},
		};
		const result = roundTrip(
			[
				["Left", shared, "Right"],
				[1, 2, 3],
			],
			[table("LeftData", "A1:B2"), table("RightData", "B1:C2")],
		);

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({ code: "invalid-input" });
		await expect(result).rejects.toThrow(/overlap/);
		expect(reads).toBe(1);
	});

	it("treats a reused object at different coordinates as distinct cell occurrences", async () => {
		let reads = 0;
		const shared: StyledCell = {
			get value() {
				return ++reads === 1 ? "First" : "Second";
			},
		};

		const sheet = await roundTrip(
			[
				[shared, shared],
				[1, 2],
			],
			[table("Data", "A1:B2")],
		);

		expect(sheet.tables[0]?.columns.map((c) => c.name)).toEqual(["First", "Second"]);
		expect(sheet.cell("A1").value).toBe("First");
		expect(sheet.cell("B1").value).toBe("Second");
		expect(reads).toBe(2);
	});

	const nonText: readonly { label: string; value: CellValue }[] = [
		{ label: "number", value: 42 },
		{ label: "boolean", value: true },
		{ label: "null", value: null },
		{ label: "undefined", value: undefined },
		{ label: "empty string", value: "" },
	];
	it.each(nonText)("rejects a captured $label header without fallback", async ({ value }) => {
		let reads = 0;
		const header: StyledCell = {
			get value() {
				return ++reads === 1 ? value : "LaterText";
			},
		};
		const result = roundTrip([[header], [42]]);

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({ code: "invalid-input" });
		await expect(result).rejects.toThrow(/must be non-empty text/);
		expect(reads).toBe(1);
	});

	it.each(
		nonText,
	)("retains the captured $label cell while using provided column-name fallback", async ({
		value,
	}) => {
		let reads = 0;
		const header: StyledCell = {
			get value() {
				return ++reads === 1 ? value : "LaterText";
			},
		};
		const sheet = await roundTrip(
			[[header], [42]],
			[{ ...table(), columns: [{ name: "Fallback" }] }],
		);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Fallback");
		expect(sheet.cell("A1").value).toBe(value ?? null);
		expect(reads).toBe(1);
	});

	it.each([
		{ label: "numeric", value: 42 },
		{ label: "uncached", value: undefined },
	])("retains $label formula-header policy", async ({ value }) => {
		const header: StyledCell = { formula: "21*2", value };
		await expect(roundTrip([[header], [42]])).rejects.toMatchObject({
			code: "invalid-input",
		});

		const sheet = await roundTrip(
			[[header], [42]],
			[{ ...table(), columns: [{ name: "Fallback" }] }],
		);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Fallback");
		expect(sheet.cell("A1").value).toBe(value ?? null);
		expect(sheet.formula("A1")).toBe("21*2");
	});

	it("keeps header text authoritative over a provided different name", async () => {
		const sheet = await roundTrip(
			[[{ value: "Header" }], [42]],
			[{ ...table(), columns: [{ name: "Provided" }] }],
		);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
	});

	it("preserves numeric and missing cells in shifted sparse header rows", async () => {
		let reads = 0;
		const header: StyledCell = {
			get value() {
				if (++reads > 1) throw new Error("shifted header read twice");
				return "Shifted";
			},
		};
		const row: CellInput[] = [];
		row[2] = header;
		row[4] = 2026;
		const sheet = await roundTrip(
			[undefined, undefined, row, [undefined, undefined, 1, 2, 3]],
			[
				{
					...table("ShiftedData", "C3:E4"),
					columns: [{ name: "Ignored" }, { name: "Missing" }, { name: "Year" }],
				},
			],
		);

		expect(sheet.tables[0]?.columns.map((c) => c.name)).toEqual(["Shifted", "Missing", "Year"]);
		expect(sheet.cell("C3").value).toBe("Shifted");
		expect(sheet.cell("D3").value).toBeNull();
		expect(sheet.cell("E3").value).toBe(2026);
		expect(reads).toBe(1);
	});

	it("does not infer names when the table has no header row", async () => {
		let reads = 0;
		const cell: StyledCell = {
			get value() {
				if (++reads > 1) throw new Error("ordinary data cell read twice");
				return "Data";
			},
		};
		const sheet = await roundTrip([[cell], [42]], [{ ...table(), headerRow: false }]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Column1");
		expect(sheet.tables[0]?.headerRow).toBe(false);
		expect(sheet.cell("A1").value).toBe("Data");
		expect(reads).toBe(1);
	});

	it("does not hide invalid cell fields while capturing a header", async () => {
		const header = { value: "Header", invalid: "must reject" };

		await expect(roundTrip([[header], [42]])).rejects.toMatchObject({
			code: "invalid-input",
		});
	});

	it("preserves the existing fallback policy for non-plain header objects", async () => {
		class Header {
			value = "Header";
		}
		await expect(roundTrip([[new Header()], [42]])).rejects.toMatchObject({
			code: "invalid-input",
		});

		const sheet = await roundTrip(
			[[new Header()], [42]],
			[{ ...table(), columns: [{ name: "Fallback" }] }],
		);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Fallback");
		expect(sheet.cell("A1").value).toBe("Header");
	});

	it("keeps null-prototype string headers eligible for inference", async () => {
		let reads = 0;
		const header: StyledCell = {
			get value() {
				if (++reads > 1) throw new Error("null-prototype header read twice");
				return "Header";
			},
		};
		Object.setPrototypeOf(header, null);

		const sheet = await roundTrip([[header], [42]]);

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Header");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(reads).toBe(1);
	});

	it("preserves frozen inputs and ordinary bytes", async () => {
		const header = Object.freeze({ value: "Header", formula: '"Header"' });
		const row = Object.freeze([header]);
		const rows = Object.freeze([row, Object.freeze([42])]);
		const info = Object.freeze({ ...table(), columns: Object.freeze([]) });
		const sheet = Object.freeze({ name: "S", rows, tables: Object.freeze([info]) });
		const workbook = Object.freeze({ sheets: Object.freeze([sheet]) });
		const before = Object.getOwnPropertyDescriptors(header);
		const control = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[{ value: "Header", formula: '"Header"' }], [42]],
					tables: [table()],
				},
			],
		});

		expect(await writeXlsx(workbook)).toEqual(control);
		expect(Object.getOwnPropertyDescriptors(header)).toEqual(before);
		expect(sheet.rows).toBe(rows);
		expect(rows[0]).toBe(row);
		expect(row[0]).toBe(header);
	});

	it("keeps streaming names provided and captures header properties only during lazy emission", async () => {
		let started = 0;
		let closed = 0;
		let columnsReads = 0;
		const reads = { value: 0, formula: 0, style: 0 };
		const header: StyledCell = {
			get value() {
				if (++reads.value > 1) throw new Error("stream header value read twice");
				return "Header";
			},
			get formula() {
				if (++reads.formula > 1) throw new Error("stream formula read twice");
				return '"Header"';
			},
			get style(): CellStyle {
				if (++reads.style > 1) throw new Error("stream style read twice");
				return { font: { bold: true } };
			},
		};
		async function* rows(): AsyncGenerator<readonly CellInput[]> {
			started++;
			try {
				yield [header];
				yield [42];
			} finally {
				closed++;
			}
		}
		const info: TableInfo = {
			...table(),
			get columns() {
				columnsReads++;
				expect(started).toBe(0);
				expect(reads).toEqual({ value: 0, formula: 0, style: 0 });
				return [{ name: "Provided" }];
			},
		};
		const stream = streamXlsx({ sheets: [{ name: "S", rows: rows(), tables: [info] }] });
		const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
		const sheet = (await openXlsx(bytes)).sheet("S");

		expect(sheet.tables[0]?.columns[0]?.name).toBe("Provided");
		expect(sheet.cell("A1").value).toBe("Header");
		expect(sheet.formula("A1")).toBe('"Header"');
		expect(sheet.style("A1")?.font?.bold).toBe(true);
		expect(reads).toEqual({ value: 1, formula: 1, style: 1 });
		expect(columnsReads).toBe(1);
		expect(started).toBe(1);
		expect(closed).toBe(1);
	});

	it("rejects missing streaming column names before starting the row source", async () => {
		let started = false;
		async function* rows(): AsyncGenerator<readonly CellInput[]> {
			started = true;
			yield ["Header"];
			yield [42];
		}
		const stream = streamXlsx({ sheets: [{ name: "S", rows: rows(), tables: [table()] }] });
		const result = new Response(stream).arrayBuffer();

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({ code: "invalid-input" });
		await expect(result).rejects.toThrow(/name is required/);
		expect(started).toBe(false);
	});
});
