import { deepStrictEqual } from "node:assert";
import { isDeepStrictEqual } from "node:util";
import { describe, expect, it } from "vitest";
import { rewriteParts } from "../../reader/__tests__/shared-part-fixtures";
import { openXlsx } from "../../reader/workbook";
import { workbookToInput } from "../from-workbook";
import { optionsDate1904, requireWorkbookObject, validateSheetMeta } from "../parts";
import { providedColumn, splitInput } from "../sheet";
import { streamXlsx } from "../stream";
import type { CellInput, SheetInput, StyledCell, WorkbookInput, WriteOptions } from "../types";
import { writeXlsx } from "../workbook";
import { MAX_PROTOTYPE_DEPTH, rejectPrototypeFields } from "../xml";

// F11.34: the writer reads only own data properties (F11.23), so a field that a class or
// Object.create prototype supplies would be dropped without an error. Such an input now fails with
// invalid-input, naming the field, without any getter running. Plain and null-prototype records keep
// F11.23's policy: a key polluted onto Object.prototype stays ignored. Bridge sheets are ordinary
// objects whose carried fields are own data properties.

const typed = (key: string) =>
	expect.objectContaining({
		name: "XlsxError",
		code: "invalid-input",
		message: expect.stringContaining(`"${key}" comes from a prototype`),
	});

async function streamed(input: WorkbookInput, options?: WriteOptions): Promise<Uint8Array> {
	const chunks: Uint8Array[] = [];
	const reader = streamXlsx(input, options).getReader();
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
	}
	const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.length;
	}
	return bytes;
}

const writers = [
	{
		name: "writeXlsx",
		write: (input: WorkbookInput, options?: WriteOptions) => writeXlsx(input, options),
	},
	{ name: "streamXlsx", write: streamed },
];

// Run with `key` defined on Object.prototype, then restore it.
async function polluted(keys: readonly string[], run: () => Promise<void>): Promise<void> {
	const previous = keys.map((key) => Object.getOwnPropertyDescriptor(Object.prototype, key));
	for (const key of keys) {
		Object.defineProperty(Object.prototype, key, {
			configurable: true,
			writable: true,
			value:
				key === "merges" ? ["A1:B2"] : key === "formula" ? "1+1" : { font: { bold: true } },
		});
	}
	try {
		await run();
	} finally {
		keys.forEach((key, i) => {
			const descriptor = previous[i];
			if (descriptor === undefined) Reflect.deleteProperty(Object.prototype, key);
			else Object.defineProperty(Object.prototype, key, descriptor);
		});
	}
}

// A prototype getter that counts its calls, so a pin can assert none ran.
function counted<T>(value: T) {
	const counter = { reads: 0 };
	const descriptor: PropertyDescriptor = {
		configurable: true,
		get() {
			counter.reads++;
			return value;
		},
	};
	return { counter, descriptor };
}

// An instance of a class whose prototype supplies `key` through a getter, with `own` fields of its own.
// The own fields are defined, not assigned, so one can shadow the getter-only prototype field.
function withPrototypeField<T extends object>(own: T, key: string, value: unknown) {
	const { counter, descriptor } = counted(value);
	const proto = {};
	Object.defineProperty(proto, key, descriptor);
	class Holder {}
	Object.setPrototypeOf(Holder.prototype, proto);
	const record: T = Object.create(Holder.prototype, Object.getOwnPropertyDescriptors(own));
	return { record, counter };
}

describe("rejectPrototypeFields", () => {
	const keys = ["value", "style", "formula"];

	it("returns for plain, null-prototype and own-field records", () => {
		rejectPrototypeFields({ value: 5 }, keys, "cell A1");
		rejectPrototypeFields(Object.create(null), keys, "cell A1");
		rejectPrototypeFields(
			new (class {
				value = 5;
				style = {};
			})(),
			keys,
			"cell A1",
		);
		// A prototype that supplies only keys outside the list.
		rejectPrototypeFields(
			new (class {
				value = 5;
				describe() {
					return "cell";
				}
			})(),
			keys,
			"cell A1",
		);
	});

	it("throws for a field on any prototype in the chain, without reading it", () => {
		const { record, counter } = withPrototypeField({ value: 5 }, "style", {});
		expect(() => rejectPrototypeFields(record, keys, "cell A1")).toThrow(typed("style"));
		expect(() => rejectPrototypeFields(record, keys, "cell A1")).toThrow(/^cell A1: /);
		class Base {
			get formula() {
				counter.reads++;
				return "1+1";
			}
		}
		class Derived extends Base {}
		expect(() => rejectPrototypeFields(new Derived(), keys, "cell B2")).toThrow(
			typed("formula"),
		);
		// A method is a prototype-supplied field too.
		expect(() =>
			rejectPrototypeFields(
				new (class {
					value = 5;
					style() {
						return {};
					}
				})(),
				keys,
				"cell A1",
			),
		).toThrow(typed("style"));
		// A chain that ends in null without passing Object.prototype.
		const base = Object.create(null);
		base.value = 5;
		expect(() =>
			rejectPrototypeFields(Object.create(Object.create(base)), keys, "cell A1"),
		).toThrow(typed("value"));
		expect(counter.reads).toBe(0);
	});

	it("stops at Object.prototype, so a key polluted there stays ignored", async () => {
		await polluted(["style", "formula"], async () => {
			rejectPrototypeFields({ value: 5 }, keys, "cell A1");
			rejectPrototypeFields(
				new (class {
					value = 5;
				})(),
				keys,
				"cell A1",
			);
		});
	});

	it("returns for a field the record owns even when a prototype has it too, without reading the prototype's", () => {
		const { record, counter } = withPrototypeField({ value: 5, style: {} }, "style", {});
		rejectPrototypeFields(record, keys, "cell A1");
		rejectPrototypeFields(
			Object.assign(Object.create({ style: { font: { italic: true } } }), {
				value: 5,
				style: {},
			}),
			keys,
			"cell A1",
		);
		expect(counter.reads).toBe(0);
	});

	it("walks a prototype that is a function", () => {
		const proto = Object.assign(() => 0, { style: {} });
		expect(() =>
			rejectPrototypeFields(Object.setPrototypeOf({ value: 5 }, proto), keys, "cell A1"),
		).toThrow(typed("style"));
		// A function's own `name` and `length` are not model fields of a cell.
		rejectPrototypeFields(
			Object.setPrototypeOf({ value: 5 }, () => 0),
			keys,
			"cell A1",
		);
	});

	it(`rejects a chain longer than ${MAX_PROTOTYPE_DEPTH}, an endless one too`, () => {
		const chain = (depth: number) => {
			let proto: object | null = null;
			for (let i = 0; i < depth; i++) proto = Object.create(proto);
			return Object.create(proto);
		};
		rejectPrototypeFields(chain(MAX_PROTOTYPE_DEPTH), keys, "cell A1");
		const tooLong = expect.objectContaining({
			name: "XlsxError",
			code: "invalid-input",
			message: `cell A1: its prototype chain is longer than ${MAX_PROTOTYPE_DEPTH}; use a plain object`,
		});
		expect(() =>
			rejectPrototypeFields(chain(MAX_PROTOTYPE_DEPTH + 1), keys, "cell A1"),
		).toThrow(tooLong);
		const endless = (): object => new Proxy({}, { getPrototypeOf: () => endless() });
		expect(() => rejectPrototypeFields(endless(), keys, "cell A1")).toThrow(tooLong);
	});

	it("sees a field that a Proxy's getPrototypeOf trap supplies", () => {
		const proto = { style: {} };
		const record = new Proxy({ value: 5 }, { getPrototypeOf: () => proto });
		expect(() => rejectPrototypeFields(record, keys, "cell A1")).toThrow(typed("style"));
	});

	it("does not see a field that a Proxy over a plain target serves from its get trap alone", () => {
		const record = new Proxy<StyledCell>(
			{ value: 5 },
			{
				get: (target, key) =>
					key === "style" ? { font: { bold: true } } : Reflect.get(target, key),
			},
		);
		expect(() => rejectPrototypeFields(record, keys, "cell A1")).not.toThrow();
		// Nor one a getPrototypeOf trap hides behind a plain prototype.
		const { record: instance } = withPrototypeField({ value: 5 }, "style", {});
		const hidden = new Proxy(instance, { getPrototypeOf: () => Object.prototype });
		expect(() => rejectPrototypeFields(hidden, keys, "cell A1")).not.toThrow();
	});
});

describe("splitInput", () => {
	it("captures empty, primitive and Date cells", () => {
		const date = new Date(Date.UTC(2020, 0, 1));
		expect(splitInput(1, 1, null)).toEqual({
			value: null,
			formula: undefined,
			style: undefined,
		});
		expect(splitInput(1, 1, undefined)).toEqual({
			value: undefined,
			formula: undefined,
			style: undefined,
		});
		expect(splitInput(1, 1, 5)).toEqual({ value: 5, formula: undefined, style: undefined });
		expect(splitInput(1, 1, date)).toEqual({
			value: date,
			formula: undefined,
			style: undefined,
		});
	});

	it("rejects an array, a prototype field, an unknown key, an empty object and an object value", () => {
		expect(() => splitInput(1, 1, [5] as never)).toThrow(
			"cell A1: an array is not a cell value",
		);
		const { record, counter } = withPrototypeField({ value: 5 }, "style", {
			font: { bold: true },
		});
		expect(() => splitInput(2, 3, record)).toThrow(typed("style"));
		expect(() => splitInput(2, 3, record)).toThrow(/^cell B3: /);
		expect(counter.reads).toBe(0);
		expect(() => splitInput(1, 1, { value: 5, typo: 1 } as never)).toThrow('(got "typo")');
		expect(() => splitInput(1, 1, {})).toThrow(
			"an object cell must be { value, style? } or carry a formula",
		);
		expect(() => splitInput(1, 1, { value: { value: 5 } } as never)).toThrow(
			"a cell's value cannot be an object",
		);
	});

	it("captures a styled or formula cell's own fields", () => {
		const style = { font: { bold: true } };
		expect(splitInput(1, 1, { formula: "1+1", value: 2, style })).toEqual({
			value: 2,
			formula: "1+1",
			style,
		});
		expect(
			splitInput(
				1,
				1,
				new (class {
					value = 5;
				})(),
			),
		).toEqual({ value: 5, formula: undefined, style: undefined });
		const { record } = withPrototypeField({ value: 5, style }, "style", {});
		expect(splitInput(1, 1, record)).toEqual({ value: 5, formula: undefined, style });
	});
});

describe("requireWorkbookObject", () => {
	it("rejects a non-object and a workbook whose prototype supplies a field", () => {
		expect(() => requireWorkbookObject(null)).toThrow("a workbook must be an object");
		expect(() => requireWorkbookObject(5)).toThrow("a workbook must be an object");
		const { record, counter } = withPrototypeField({ sheets: [] }, "protection", {});
		expect(() => requireWorkbookObject(record)).toThrow(typed("protection"));
		expect(() => requireWorkbookObject(record)).toThrow(/^workbook: /);
		expect(counter.reads).toBe(0);
	});

	it("returns for a plain, null-prototype or own-field workbook", () => {
		requireWorkbookObject({ sheets: [] });
		requireWorkbookObject(Object.create(null));
		requireWorkbookObject(
			new (class {
				sheets = [];
			})(),
		);
	});
});

describe("validateSheetMeta", () => {
	const sheet = (name: string, extra: Partial<SheetInput> = {}): SheetInput => ({
		name,
		rows: [],
		...extra,
	});

	it("rejects an empty list and a sheet whose prototype supplies a field", () => {
		expect(() => validateSheetMeta([])).toThrow("a workbook needs at least one sheet");
		const { record, counter } = withPrototypeField(sheet("S"), "merges", ["A1:B2"]);
		expect(() => validateSheetMeta([sheet("A"), record])).toThrow(typed("merges"));
		expect(() => validateSheetMeta([sheet("A"), record])).toThrow(/^sheets\[1\]: /);
		expect(counter.reads).toBe(0);
	});

	it("rejects a bad, long, forbidden, unsafe or duplicate name", () => {
		expect(() => validateSheetMeta([sheet("")])).toThrow(
			"a sheet name must be a non-empty string",
		);
		expect(() => validateSheetMeta([sheet("x".repeat(32))])).toThrow("exceeds 31 characters");
		expect(() => validateSheetMeta([sheet("a/b")])).toThrow("contains a forbidden character");
		expect(() => validateSheetMeta([sheet("a\u0001")])).toThrow("not allowed in XML");
		expect(() => validateSheetMeta([sheet("S"), sheet("s")])).toThrow(
			'duplicate sheet name "s"',
		);
	});

	it("rejects a bad state and a workbook with no visible sheet", () => {
		expect(() => validateSheetMeta([sheet("S", { state: "gone" as never })])).toThrow(
			'sheet "S": state must be',
		);
		expect(() => validateSheetMeta([sheet("S", { state: "hidden" })])).toThrow(
			"at least one sheet must be visible",
		);
	});

	it("returns the names and states, for class-instance sheets with own fields too", () => {
		class Sheet {
			name = "B";
			rows = [];
			state = "hidden" as const;
		}
		expect(validateSheetMeta([sheet("A"), new Sheet()])).toEqual({
			states: ["visible", "hidden"],
			names: ["A", "B"],
		});
	});
});

describe("providedColumn", () => {
	const notAnObject = (what: string) =>
		expect.objectContaining({
			name: "XlsxError",
			code: "invalid-input",
			message: `sheet "S": ${what} must be an object`,
		});

	it("is undefined for an absent entry: no columns, a shorter array or a hole", () => {
		expect(providedColumn("S", "tables[0]", undefined, 0)).toBeUndefined();
		expect(providedColumn("S", "tables[0]", [], 0)).toBeUndefined();
		const holes: unknown[] = [];
		holes.length = 2;
		expect(providedColumn("S", "tables[0]", holes, 1)).toBeUndefined();
	});

	it("returns a plain or null-prototype entry", () => {
		const plain = { name: "Amount" };
		expect(providedColumn("S", "tables[0]", [plain], 0)).toBe(plain);
		const bare = Object.assign(Object.create(null), { name: "Amount" });
		expect(providedColumn("S", "tables[0]", [bare], 0)).toBe(bare);
	});

	it("rejects any other entry, typed", () => {
		const entries: unknown[] = [
			new (class {
				name = "Amount";
			})(),
			Object.create({ name: "Amount" }),
			"Amount",
			null,
			5,
			["Amount"],
		];
		for (const entry of entries) {
			expect(() => providedColumn("S", "tables[2]", [{}, entry], 1)).toThrow(
				notAnObject("tables[2].columns[1]"),
			);
		}
	});

	it("reads the entry once", () => {
		let reads = 0;
		const columns: unknown[] = [];
		Object.defineProperty(columns, 0, {
			get() {
				reads++;
				return { name: "Amount" };
			},
		});
		expect(providedColumn("S", "tables[0]", columns, 0)).toEqual({ name: "Amount" });
		expect(reads).toBe(1);
	});
});

describe("optionsDate1904", () => {
	it("is undefined without options or an own flag", async () => {
		expect(optionsDate1904(undefined)).toBeUndefined();
		expect(optionsDate1904(null)).toBeUndefined();
		expect(optionsDate1904({})).toBeUndefined();
		await polluted(["date1904"], async () => {
			expect(optionsDate1904({})).toBeUndefined();
		});
	});

	it("returns an own flag on a plain, null-prototype or class-instance object", () => {
		expect(optionsDate1904({ date1904: true })).toBe(true);
		expect(optionsDate1904({ date1904: false })).toBe(false);
		const bare: WriteOptions = Object.assign(Object.create(null), { date1904: true });
		expect(optionsDate1904(bare)).toBe(true);
		expect(
			optionsDate1904(
				new (class {
					readonly date1904 = true;
				})(),
			),
		).toBe(true);
	});

	it("rejects a flag that a prototype supplies, without reading it", () => {
		const { record, counter } = withPrototypeField({}, "date1904", true);
		expect(() => optionsDate1904(record)).toThrow(typed("date1904"));
		expect(() => optionsDate1904(record)).toThrow(/^options: /);
		const created: WriteOptions = Object.create({ date1904: true });
		expect(() => optionsDate1904(created)).toThrow(typed("date1904"));
		expect(counter.reads).toBe(0);
	});
});

for (const { name, write } of writers) {
	describe(`${name} with prototype-supplied fields`, () => {
		const plainBytes = () => write({ sheets: [{ name: "S", rows: [[5]] }] });

		it.each([
			["style", { font: { bold: true } }],
			["formula", "1+1"],
		])("rejects a cell whose prototype supplies %s", async (key, value) => {
			const { record, counter } = withPrototypeField({ value: 5 }, key, value);
			await expect(write({ sheets: [{ name: "S", rows: [[record]] }] })).rejects.toThrow(
				typed(key),
			);
			expect(counter.reads).toBe(0);
		});

		it.each([
			["merges", ["A1:B2"]],
			["freeze", { rows: 1 }],
			["columns", [{ min: 1, max: 1, width: 20 }]],
		])("rejects a sheet whose prototype supplies %s", async (key, value) => {
			const { record, counter } = withPrototypeField({ name: "S", rows: [[5]] }, key, value);
			await expect(write({ sheets: [record] })).rejects.toThrow(typed(key));
			expect(counter.reads).toBe(0);
		});

		it.each([
			["definedNames", [{ name: "Total", refersTo: "S!$A$1" }]],
			["protection", { lockStructure: true }],
		])("rejects a workbook whose prototype supplies %s", async (key, value) => {
			const { record, counter } = withPrototypeField(
				{ sheets: [{ name: "S", rows: [[5]] }] },
				key,
				value,
			);
			await expect(write(record)).rejects.toThrow(typed(key));
			expect(counter.reads).toBe(0);
		});

		it("rejects Object.create inputs whose prototype supplies a field", async () => {
			const cell: CellInput = Object.create(
				{ style: { font: { bold: true } } },
				{
					value: { value: 5, enumerable: true },
				},
			);
			await expect(write({ sheets: [{ name: "S", rows: [[cell]] }] })).rejects.toThrow(
				typed("style"),
			);
			const sheet: SheetInput = Object.create(
				{ merges: ["A1:B2"] },
				{
					name: { value: "S", enumerable: true },
					rows: { value: [[5]], enumerable: true },
				},
			);
			await expect(write({ sheets: [sheet] })).rejects.toThrow(typed("merges"));
			const book: WorkbookInput = Object.create(
				{ protection: { lockStructure: true } },
				{
					sheets: { value: [{ name: "S", rows: [[5]] }], enumerable: true },
				},
			);
			await expect(write(book)).rejects.toThrow(typed("protection"));
		});

		it("writes an own field that shadows a prototype one, without reading the prototype's", async () => {
			const bold = { font: { bold: true } };
			const control = await write({
				sheets: [{ name: "S", rows: [[{ value: 5, style: bold }]], merges: ["A2:B2"] }],
			});
			const { record: cell, counter: cellReads } = withPrototypeField(
				{ value: 5, style: bold },
				"style",
				{ font: { italic: true } },
			);
			const { record: sheet, counter: sheetReads } = withPrototypeField(
				{ name: "S", rows: [[cell]], merges: ["A2:B2"] },
				"merges",
				["A1:B1"],
			);
			expect(await write({ sheets: [sheet] })).toEqual(control);
			const created: CellInput = Object.assign(
				Object.create({ style: { font: { italic: true } } }),
				{
					value: 5,
					style: bold,
				},
			);
			expect(
				await write({ sheets: [{ name: "S", rows: [[created]], merges: ["A2:B2"] }] }),
			).toEqual(control);
			expect([cellReads.reads, sheetReads.reads]).toEqual([0, 0]);
		});

		it("checks a sheet that is a function, such as a class with static fields", async () => {
			let reads = 0;
			// biome-ignore lint/complexity/noStaticOnlyClass: a class used as a sheet is the input under test
			class Base {
				static get merges() {
					reads++;
					return ["A1:B1"];
				}
			}
			class Merged extends Base {
				static rows = [[5]];
			}
			await expect(write({ sheets: [Merged] })).rejects.toThrow(typed("merges"));
			expect(reads).toBe(0);
			// A function's own `name` names the sheet, as before.
			// biome-ignore lint/complexity/noStaticOnlyClass: a class used as a sheet is the input under test
			class S {
				static rows = [[5]];
			}
			expect(await write({ sheets: [S] })).toEqual(await plainBytes());
		});

		it("writes class-instance cells and sheets with own fields unchanged, Object.prototype polluted or not", async () => {
			const control = await plainBytes();
			class Cell {
				readonly value: number;
				constructor() {
					this.value = 5;
				}
			}
			class Sheet {
				readonly name = "S";
				readonly rows = [[new Cell()]];
			}
			class Book {
				readonly sheets = [new Sheet()];
			}
			expect(await write(new Book())).toEqual(control);
			await polluted(["style", "formula", "merges", "definedNames"], async () => {
				expect(await write({ sheets: [new Sheet()] })).toEqual(control);
				expect(await write(new Book())).toEqual(control);
				expect(await write({ sheets: [{ name: "S", rows: [[5]] }] })).toEqual(control);
			});
		});

		it.each([
			[
				"a class instance",
				() =>
					new (class {
						name = "Amount";
						totalsRowFunction = "sum";
					})(),
			],
			["a string", () => "Amount"],
			["null", () => null],
		])("rejects %s as a table column entry", async (_label, entry) => {
			// A JS caller can pass a column entry the types forbid.
			const table = {
				name: "T",
				ref: "A1:A3",
				headerRow: true,
				totalsRow: true,
				columns: [entry()] as never,
			};
			await expect(
				write({
					sheets: [{ name: "S", rows: [["Amount"], [1], [null]], tables: [table] }],
				}),
			).rejects.toThrow(
				expect.objectContaining({
					code: "invalid-input",
					message: 'sheet "S": tables[0].columns[0] must be an object',
				}),
			);
		});

		it("writes a plain table column entry's fields", async () => {
			const bytes = await write({
				sheets: [
					{
						name: "S",
						rows: [["Amount"], [1], [null]],
						tables: [
							{
								name: "T",
								ref: "A1:A3",
								headerRow: true,
								totalsRow: true,
								columns: [{ name: "Amount", totalsRowFunction: "sum" }],
							},
						],
					},
				],
			});
			const table = (await openXlsx(bytes)).sheet("S").tables[0];
			expect(table?.columns?.[0]?.totalsRowFunction).toBe("sum");
		});

		it("rejects options whose prototype supplies date1904, without running the getter", async () => {
			const input: WorkbookInput = { sheets: [{ name: "S", rows: [[5]] }] };
			const { record, counter } = withPrototypeField({}, "date1904", true);
			await expect(write(input, record)).rejects.toThrow(typed("date1904"));
			const created: WriteOptions = Object.create({ date1904: true });
			await expect(write(input, created)).rejects.toThrow(typed("date1904"));
			expect(counter.reads).toBe(0);
		});

		it("reads own, class-instance and null-prototype options as before", async () => {
			const input: WorkbookInput = {
				sheets: [{ name: "S", rows: [[new Date(Date.UTC(2020, 0, 1))]] }],
			};
			const control = await write(input, { date1904: true });
			expect((await openXlsx(control)).date1904).toBe(true);
			class Options {
				readonly date1904 = true;
			}
			expect(await write(input, new Options())).toEqual(control);
			const bare: WriteOptions = Object.assign(Object.create(null), { date1904: true });
			expect(await write(input, bare)).toEqual(control);
		});

		it("keeps the get-trap-only Proxy limitation: a field served from the trap alone is not written", async () => {
			const control = await plainBytes();
			const cell = new Proxy<StyledCell>(
				{ value: 5 },
				{
					get: (target, key) =>
						key === "style" ? { font: { bold: true } } : Reflect.get(target, key),
				},
			);
			expect(await write({ sheets: [{ name: "S", rows: [[cell]] }] })).toEqual(control);
			const sheet = new Proxy<SheetInput>(
				{ name: "S", rows: [[5]] },
				{ get: (target, key) => (key === "merges" ? ["A1:B2"] : Reflect.get(target, key)) },
			);
			expect(await write({ sheets: [sheet] })).toEqual(control);
		});
	});
}

describe("workbookToInput", () => {
	async function book(input: WorkbookInput) {
		return openXlsx(await writeXlsx(input));
	}

	it("returns ordinary sheet objects that compare equal to plain literals", async () => {
		const out = await workbookToInput(
			await book({ sheets: [{ name: "S", rows: [[1, "a"]] }] }),
		);
		const sheet = out.sheets[0];
		expect(Object.getPrototypeOf(sheet)).toBe(Object.prototype);
		expect(String(sheet)).toBe("[object Object]");
		expect(typeof sheet?.hasOwnProperty).toBe("function");
		deepStrictEqual(out, { sheets: [{ name: "S", rows: [[1, "a"]] }] });
		expect(isDeepStrictEqual(sheet, { ...sheet })).toBe(true);
		expect(out).toStrictEqual({ sheets: [{ name: "S", rows: [[1, "a"]] }] });
	});

	it("carries every field as an own property of an ordinary object", async () => {
		const out = await workbookToInput(
			await book({
				sheets: [
					{
						name: "S",
						rows: [
							[1, 2],
							[3, 4],
						],
						columns: [{ min: 1, max: 1, width: 20 }],
						freeze: { rows: 1 },
						merges: ["A1:B1"],
						state: "visible",
						comments: [{ ref: "A1", text: "Note" }],
						autoFilter: { ref: "A1:B2" },
						protection: { sheet: true },
						pageMargins: {
							left: 1,
							right: 1,
							top: 1,
							bottom: 1,
							header: 0.5,
							footer: 0.5,
						},
					},
					{ name: "H", rows: [[1]], state: "hidden" },
				],
			}),
		);
		const [sheet, hidden] = out.sheets;
		expect(Object.getPrototypeOf(sheet)).toBe(Object.prototype);
		for (const key of [
			"columns",
			"freeze",
			"merges",
			"comments",
			"autoFilter",
			"protection",
			"pageMargins",
		]) {
			expect([key, Object.hasOwn(sheet ?? {}, key)]).toEqual([key, true]);
		}
		expect(Object.hasOwn(sheet ?? {}, "state")).toBe(false);
		expect(hidden?.state).toBe("hidden");
		expect(sheet?.merges).toEqual(["A1:B1"]);
	});

	it("never calls an inherited Object.prototype setter for a carried field", async () => {
		const source = await book({ sheets: [{ name: "S", rows: [[1, 2]], merges: ["A1:B1"] }] });
		let calls = 0;
		const previous = Object.getOwnPropertyDescriptor(Object.prototype, "merges");
		Object.defineProperty(Object.prototype, "merges", {
			configurable: true,
			get: () => undefined,
			set: () => {
				calls++;
			},
		});
		let out: WorkbookInput | undefined;
		try {
			out = await workbookToInput(source);
		} finally {
			if (previous === undefined) Reflect.deleteProperty(Object.prototype, "merges");
			else Object.defineProperty(Object.prototype, "merges", previous);
		}
		expect(calls).toBe(0);
		expect(Object.hasOwn(out?.sheets[0] ?? {}, "merges")).toBe(true);
		expect(out?.sheets[0]?.merges).toEqual(["A1:B1"]);
	});

	it("rejects a cell with no grid position or two cells on one position", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1, 2]] }] });
		const sheetXml = "xl/worksheets/sheet1.xml";
		const offGrid = await rewriteParts(bytes, {
			[sheetXml]: (xml) => xml.replace('<c r="A1"', '<c r="A99999999999999"'),
		});
		await expect(workbookToInput(await openXlsx(offGrid))).rejects.toThrow(
			expect.objectContaining({
				code: "invalid-input",
				message: expect.stringContaining("has no writable grid position"),
			}),
		);
		const twin = await rewriteParts(bytes, {
			[sheetXml]: (xml) => xml.replace('r="B1"', 'r="a1"'),
		});
		await expect(workbookToInput(await openXlsx(twin))).rejects.toThrow(
			expect.objectContaining({
				code: "invalid-input",
				message: expect.stringContaining('"A1" and "a1"'),
			}),
		);
	});
});
