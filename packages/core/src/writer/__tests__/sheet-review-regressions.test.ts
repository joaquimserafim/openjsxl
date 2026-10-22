import { describe, expect, it } from "vitest";
import { createMediaRegistry } from "../images";
import { createTableContext, streamWorksheet, worksheetXml } from "../sheet";
import { createStyleRegistry } from "../styles";
import type { CellInput, SheetInput, StreamSheetInput } from "../types";
import { writeXlsx } from "../workbook";

function render(sheet: SheetInput) {
	return worksheetXml(
		sheet,
		sheet.rows,
		0,
		false,
		createStyleRegistry(),
		createMediaRegistry(),
		createTableContext(),
		"Data",
	);
}

async function stream(rows: StreamSheetInput["rows"]): Promise<string> {
	const result = streamWorksheet(
		{ name: "Data" },
		rows,
		0,
		false,
		createStyleRegistry(),
		createMediaRegistry(),
		createTableContext(),
		"Data",
	);
	let xml = "";
	for await (const chunk of result.chunks) xml += new TextDecoder().decode(chunk);
	return xml;
}

describe("worksheet own-property review regressions", () => {
	it.each(["native", "proxy"] as const)(
		"rejects the first absent sheet in a 10M-slot %s array without reading later entries",
		async (kind) => {
			const input: SheetInput[] = kind === "native" ? new Array(10_000_000) : [];
			let laterReads = 0;
			let lengthReads = 0;
			Object.defineProperty(input, "9999999", {
				get() {
					laterReads++;
					throw new Error("later sheet must not be read");
				},
			});
			const sheets =
				kind === "proxy"
					? new Proxy(input, {
							get(target, key, receiver) {
								if (key === "length") {
									lengthReads++;
									return 10_000_000;
								}
								return Reflect.get(target, key, receiver);
							},
						})
					: input;
			await expect(writeXlsx({ sheets })).rejects.toMatchObject({ code: "invalid-input" });
			expect(laterReads).toBe(0);
			if (kind === "proxy") expect(lengthReads).toBe(1);
		},
	);

	it.each(["sync", "async"] as const)(
		"keeps a custom array %s iterator and reads its getter once",
		async (kind) => {
			const rows = [[1]];
			let reads = 0;
			const iterator = kind === "sync" ? Symbol.iterator : Symbol.asyncIterator;
			Object.defineProperty(rows, iterator, {
				get() {
					if (++reads > 1) throw new Error("iterator read twice");
					return kind === "sync"
						? function* () {
								yield [99];
							}
						: async function* () {
								yield [99];
							};
				},
			});
			const xml = await stream(rows);
			expect(reads).toBe(1);
			expect(xml).toContain('<c r="A1"><v>99</v></c>');
			expect(xml).not.toContain("<v>1</v>");
		},
	);

	it("keeps custom iteration on an array subclass", async () => {
		class Rows extends Array<readonly CellInput[]> {
			override *[Symbol.iterator](): ArrayIterator<readonly CellInput[]> {
				yield [77];
			}
		}
		const rows = new Rows([1]);
		expect(await stream(rows)).toContain('<c r="A1"><v>77</v></c>');
	});

	it("keeps native sparse row and cell holes empty in both writers", async () => {
		const firstRow: CellInput[] = [1];
		firstRow[2] = 3;
		const rows: (CellInput[] | undefined)[] = [firstRow];
		rows[2] = [4];
		const expectedBuffered = render({ name: "Data", rows });
		const expectedStream = await stream(rows);
		const previous = Object.getOwnPropertyDescriptor(Array.prototype, "1");
		let actualBuffered: ReturnType<typeof render>;
		let actualStream: string;
		Object.defineProperty(Array.prototype, "1", {
			configurable: true,
			writable: true,
			value: [55],
		});
		try {
			actualBuffered = render({ name: "Data", rows });
			actualStream = await stream(rows);
		} finally {
			if (previous === undefined) Reflect.deleteProperty(Array.prototype, "1");
			else Object.defineProperty(Array.prototype, "1", previous);
		}
		expect(actualBuffered).toEqual(expectedBuffered);
		expect(actualStream).toBe(expectedStream);
	});

	it("captures a table header row and its cell getters once for metadata and emission", () => {
		let rowReads = 0;
		let leftReads = 0;
		let rightReads = 0;
		const headers: CellInput[] = [];
		Object.defineProperty(headers, "0", {
			get() {
				if (++leftReads > 1) throw new Error("left header read twice");
				return "Left";
			},
		});
		Object.defineProperty(headers, "1", {
			get() {
				if (++rightReads > 1) throw new Error("right header read twice");
				return "Right";
			},
		});
		const rows: (readonly CellInput[] | undefined)[] = [undefined, [1, 2]];
		Object.defineProperty(rows, "0", {
			get() {
				if (++rowReads > 1) throw new Error("header row read twice");
				return headers;
			},
		});
		const result = render({
			name: "Data",
			rows,
			tables: [
				{ name: "Sales", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false },
			],
		});
		expect([rowReads, leftReads, rightReads]).toEqual([1, 1, 1]);
		expect(result.xml).toContain('<c r="A1" t="inlineStr"><is><t>Left</t></is></c>');
		expect(result.xml).toContain('<c r="B1" t="inlineStr"><is><t>Right</t></is></c>');
		expect(result.tables?.[0]?.xml).toContain('<tableColumn id="1" name="Left"/>');
		expect(result.tables?.[0]?.xml).toContain('<tableColumn id="2" name="Right"/>');
	});

	it.each([
		"columns",
		"rowProperties",
		"freeze",
		"autoFilter",
		"protection",
		"printOptions",
		"pageMargins",
		"pageSetup",
		"headerFooter",
		"merges",
		"conditionalFormatting",
		"dataValidations",
		"hyperlinks",
		"comments",
		"images",
		"tables",
	])("does not invoke an inherited top-level %s getter", (key) => {
		const sheet: SheetInput = { name: "Data", rows: [[1]] };
		const expected = render(sheet);
		const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
		let reads = 0;
		let result: ReturnType<typeof render>;
		Object.defineProperty(Object.prototype, key, {
			configurable: true,
			get() {
				reads++;
				throw new Error("inherited model getter");
			},
		});
		try {
			result = render(sheet);
		} finally {
			if (previous === undefined) Reflect.deleteProperty(Object.prototype, key);
			else Object.defineProperty(Object.prototype, key, previous);
		}
		expect(reads).toBe(0);
		expect(result).toEqual(expected);
	});
});
