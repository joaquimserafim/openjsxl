import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { CellInput, SheetInput, WorkbookInput, WriteOptions } from "../types";
import { writeXlsx } from "../workbook";

async function streamed(input: WorkbookInput, options?: WriteOptions): Promise<Uint8Array> {
	const chunks: Uint8Array[] = [];
	let length = 0;
	const reader = streamXlsx(input, options).getReader();
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		length += value.length;
	}
	const bytes = new Uint8Array(length);
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.length;
	}
	return bytes;
}

async function withProperty(
	target: object,
	key: PropertyKey,
	descriptor: PropertyDescriptor,
	run: () => Promise<void>,
): Promise<void> {
	const previous = Object.getOwnPropertyDescriptor(target, key);
	Object.defineProperty(target, key, { configurable: true, ...descriptor });
	try {
		await run();
	} finally {
		if (previous === undefined) Reflect.deleteProperty(target, key);
		else Object.defineProperty(target, key, previous);
	}
}

const writers = [
	{ name: "buffered", write: writeXlsx },
	{ name: "streamed", write: streamed },
];

for (const { name, write } of writers) {
	describe(`${name} workbook own-property boundaries`, () => {
		it("requires own sheets without invoking an inherited getter", async () => {
			const input: WorkbookInput = { sheets: [] };
			Reflect.deleteProperty(input, "sheets");
			let reads = 0;
			await withProperty(
				Object.prototype,
				"sheets",
				{
					get() {
						reads++;
						return [{ name: "Injected", rows: [[42]] }];
					},
				},
				async () => {
					await expect(write(input)).rejects.toMatchObject({ code: "invalid-input" });
				},
			);
			expect(reads).toBe(0);
		});

		it("requires own rows without invoking an inherited getter", async () => {
			const sheet: SheetInput = { name: "S", rows: [] };
			Reflect.deleteProperty(sheet, "rows");
			let reads = 0;
			await withProperty(
				Object.prototype,
				"rows",
				{
					get() {
						reads++;
						return [[42]];
					},
				},
				async () => {
					await expect(write({ sheets: [sheet] })).rejects.toMatchObject({
						code: "invalid-input",
					});
				},
			);
			expect(reads).toBe(0);
		});

		it("ignores inherited date1904 on both the input and options", async () => {
			const ordinary: WorkbookInput = { sheets: [{ name: "S", rows: [[1]] }] };
			const control = await write(ordinary);
			await withProperty(
				Object.prototype,
				"date1904",
				{ value: true, writable: true },
				async () => {
					expect(await write(ordinary)).toEqual(control);
					expect(await write({ ...ordinary, date1904: false }, {})).toEqual(control);
				},
			);
		});

		it.each(["definedNames", "protection", "themeXml"])(
			"does not read inherited optional %s",
			async (key) => {
				const input: WorkbookInput = { sheets: [{ name: "S", rows: [[1]] }] };
				const control = await write(input);
				let reads = 0;
				await withProperty(
					Object.prototype,
					key,
					{
						get() {
							reads++;
							throw new Error(`inherited ${key} read`);
						},
					},
					async () => {
						expect(await write(input)).toEqual(control);
					},
				);
				expect(reads).toBe(0);
			},
		);

		it("rejects an inherited sheet array slot without reading it", async () => {
			const sheets: SheetInput[] = new Array(1);
			const prototype = Object.create(Array.prototype);
			let reads = 0;
			Object.defineProperty(prototype, "0", {
				get() {
					reads++;
					return { name: "Injected", rows: [[42]] };
				},
			});
			Object.setPrototypeOf(sheets, prototype);
			await expect(write({ sheets })).rejects.toMatchObject({ code: "invalid-input" });
			expect(reads).toBe(0);
		});

		it("captures each own sheet array slot once before validation and emission", async () => {
			let reads = 0;
			const sheets: SheetInput[] = [];
			Object.defineProperty(sheets, "0", {
				get() {
					reads++;
					if (reads > 1) throw new Error("sheet index was read again");
					return { name: "S", rows: [[42]] };
				},
			});
			const workbook = await openXlsx(await write({ sheets }));
			expect(workbook.sheet("S").cell("A1").value).toBe(42);
			expect(reads).toBe(1);
		});

		it("does not read a later sheet getter after an earlier sheet fails validation", async () => {
			const sheets: SheetInput[] = [{ name: "", rows: [] }];
			let reads = 0;
			Object.defineProperty(sheets, "1", {
				get() {
					reads++;
					throw new Error("later sheet must not be read");
				},
			});
			await expect(write({ sheets })).rejects.toMatchObject({ code: "invalid-input" });
			expect(reads).toBe(0);
		});

		it("rejects a huge sparse sheet list at its first absent item", async () => {
			const sheets: SheetInput[] = new Array(1_000_000_000);
			await expect(write({ sheets })).rejects.toMatchObject({ code: "invalid-input" });
		});

		it("reads the caller array length only once", async () => {
			let reads = 0;
			const sheets = new Proxy<SheetInput[]>([{ name: "S", rows: [[42]] }], {
				get(target, key, receiver) {
					if (key === "length") {
						reads++;
						if (reads > 1) throw new Error("array length was read again");
					}
					return Reflect.get(target, key, receiver);
				},
			});
			await write({ sheets });
			expect(reads).toBe(1);
		});

		it.each([1.5, -1, Number.NaN, Number.POSITIVE_INFINITY, "1"])(
			"rejects invalid proxy array length %s with a typed error",
			async (length) => {
				const sheets = new Proxy<SheetInput[]>([{ name: "S", rows: [] }], {
					get(target, key, receiver) {
						return key === "length" ? length : Reflect.get(target, key, receiver);
					},
				});
				await expect(write({ sheets })).rejects.toMatchObject({ code: "invalid-input" });
			},
		);

		it("reads own workbook and option proxy fields once", async () => {
			const reads = new Map<PropertyKey, number>();
			const input = new Proxy<WorkbookInput>(
				{
					sheets: [{ name: "S", rows: [[42]] }],
					date1904: true,
					themeXml: "<ignored/>",
					definedNames: [],
					protection: {},
				},
				{
					get(target, key, receiver) {
						reads.set(key, (reads.get(key) ?? 0) + 1);
						return Reflect.get(target, key, receiver);
					},
				},
			);
			let optionReads = 0;
			const options: WriteOptions = {
				get date1904() {
					optionReads++;
					return false;
				},
			};
			const bytes = await write(input, options);
			const xml = new TextDecoder().decode(await openZip(bytes).read("xl/workbook.xml"));
			expect(xml).not.toContain("date1904");
			for (const key of ["sheets", "date1904", "themeXml", "definedNames", "protection"]) {
				expect(reads.get(key)).toBe(1);
			}
			expect(optionReads).toBe(1);
		});

		it("keeps ordinary null-prototype workbook, sheet, and options bytes", async () => {
			const control = await write(
				{ sheets: [{ name: "S", rows: [[42]] }] },
				{ date1904: false },
			);
			const sheet: SheetInput = { name: "S", rows: [[42]] };
			Object.setPrototypeOf(sheet, null);
			const input: WorkbookInput = { sheets: [sheet] };
			Object.setPrototypeOf(input, null);
			const options: WriteOptions = { date1904: false };
			Object.setPrototypeOf(options, null);
			expect(await write(input, options)).toEqual(control);
		});
	});
}

it("the bridge does not reuse or mutate an inherited sparse accumulator row", async () => {
	const rows: CellInput[][] = [];
	rows[1000] = [42];
	const source = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
	const inherited: CellInput[] = [];
	let input: WorkbookInput | undefined;
	await withProperty(Array.prototype, "1000", { value: inherited, writable: true }, async () => {
		input = await workbookToInput(source);
	});
	expect(inherited).toEqual([]);
	expect(input?.sheets[0]?.rows[1000]).toEqual([42]);
	expect(Object.hasOwn(input?.sheets[0]?.rows ?? [], 1000)).toBe(true);
});

it("the bridge preserves distinct global and local names when the prototype supplies a scope", async () => {
	const source = await openXlsx(
		await writeXlsx({
			sheets: [{ name: "S", rows: [[1]] }],
			definedNames: [
				{ name: "Same_Name", refersTo: "1" },
				{ name: "Same_Name", refersTo: "2", localSheetId: 0 },
			],
		}),
	);
	let input: WorkbookInput | undefined;
	await withProperty(Object.prototype, "localSheetId", { value: 0, writable: true }, async () => {
		input = await workbookToInput(source);
	});
	expect(input?.definedNames).toEqual(source.definedNames);
	expect(input?.definedNames).toHaveLength(2);
});

it("the bridge writes own metadata despite inherited getter-only properties", async () => {
	const source = await openXlsx(
		await writeXlsx({
			sheets: [{ name: "S", rows: [[1]], comments: [{ ref: "A1", text: "Own" }] }],
		}),
	);
	let reads = 0;
	let input: WorkbookInput | undefined;
	await withProperty(
		Object.prototype,
		"comments",
		{
			get() {
				reads++;
				throw new Error("inherited comments getter");
			},
		},
		async () => {
			input = await workbookToInput(source);
		},
	);
	expect(input?.sheets[0]?.comments).toEqual([{ ref: "A1", text: "Own" }]);
	expect(reads).toBe(0);
});

it.each([
	{ axis: "row", accessor: true },
	{ axis: "row", accessor: false },
	{ axis: "column", accessor: true },
	{ axis: "column", accessor: false },
])(
	"the bridge creates own $axis slots over inherited readonly indices (accessor=$accessor)",
	async ({ axis, accessor }) => {
		const rows: CellInput[][] = [];
		const rowIndex = axis === "row" ? 1000 : 0;
		const colIndex = axis === "column" ? 1000 : 0;
		const cells: CellInput[] = [];
		cells[colIndex] = 42;
		rows[rowIndex] = cells;
		const source = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		let reads = 0;
		const descriptor: PropertyDescriptor = accessor
			? {
					get() {
						reads++;
						return undefined;
					},
				}
			: { value: undefined, writable: false };
		let input: WorkbookInput | undefined;
		await withProperty(Array.prototype, "1000", descriptor, async () => {
			input = await workbookToInput(source);
		});
		expect(input?.sheets[0]?.rows[rowIndex]?.[colIndex]).toBe(42);
		expect(reads).toBe(0);
	},
);
