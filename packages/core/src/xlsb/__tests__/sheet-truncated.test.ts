import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import * as R from "../records";
import { parseXlsbSheet, valueLength } from "../sheet";

// F11.100: a cell record cut before the end of its value yields no value. RecordData reads 0, `false`,
// `""` or a partial string past the end of a record, so a cut record used to invent a value (and the
// bridge wrote it into the xlsx). A cut record now reads like a blank of its form: its coordinates and
// complete style as on HEAD, no value, and an earlier value at the same ref survives.

const u32 = (value: number): number[] => [
	value & 0xff,
	(value >>> 8) & 0xff,
	(value >>> 16) & 0xff,
	(value >>> 24) & 0xff,
];
function record(id: number, payload: readonly number[]): number[] {
	const length: number[] = [];
	let remaining = payload.length;
	do {
		const byte = remaining & 0x7f;
		remaining >>>= 7;
		length.push(byte | (remaining === 0 ? 0 : 0x80));
	} while (remaining !== 0);
	return [...(id < 0x80 ? [id] : [id & 0xff, id >>> 8]), ...length, ...payload];
}
function real(value: number): number[] {
	const bytes = new Uint8Array(8);
	new DataView(bytes.buffer).setFloat64(0, value, true);
	return [...bytes];
}
function wide(value: string): number[] {
	const out = u32(value.length);
	for (let i = 0; i < value.length; i++) {
		const code = value.charCodeAt(i);
		out.push(code & 0xff, code >>> 8);
	}
	return out;
}
const row = (index: number): number[] => record(R.ROW, u32(index));
const full = (id: number, col: number, payload: readonly number[], style = 0): number[] =>
	record(id, [...u32(col), ...u32(style), ...payload]);
const short = (id: number, payload: readonly number[], style = 0): number[] =>
	record(id, [...u32(style), ...payload]);
const parse = (...records: readonly number[][]) =>
	parseXlsbSheet(Uint8Array.from(records.flat()), ["shared", "second"], undefined, false);

// Each form's value bytes, as decodeCell reads them.
const FULL_FORMS = [
	{ name: "BrtCellRk", id: R.NUM, value: u32((17 << 2) | 2), cell: 17 },
	{ name: "BrtCellError", id: R.BOOLERR, value: [7], cell: "#DIV/0!" },
	{ name: "BrtCellBool", id: R.BOOL, value: [1], cell: true },
	{ name: "BrtCellReal", id: R.FLOAT, value: real(2.75), cell: 2.75 },
	{ name: "BrtCellSt", id: R.INLINE_STRING, value: wide("hi"), cell: "hi" },
	{ name: "BrtCellIsst", id: R.STRING, value: u32(1), cell: "second" },
	{ name: "BrtFmlaString", id: R.FORMULA_STRING, value: wide("cached"), cell: "cached" },
	{ name: "BrtFmlaNum", id: R.FORMULA_FLOAT, value: real(42), cell: 42 },
	{ name: "BrtFmlaBool", id: R.FORMULA_BOOL, value: [1], cell: true },
	{ name: "BrtFmlaError", id: R.FORMULA_BOOLERR, value: [42], cell: "#N/A" },
];
const SHORT_FORMS = [
	{ name: "BrtShortRk", id: R.SHORT_NUM, value: u32((17 << 2) | 2), cell: 17 },
	{ name: "BrtShortError", id: R.SHORT_BOOLERR, value: [7], cell: "#DIV/0!" },
	{ name: "BrtShortBool", id: R.SHORT_BOOL, value: [1], cell: true },
	{ name: "BrtShortReal", id: R.SHORT_FLOAT, value: real(2.75), cell: 2.75 },
	{ name: "BrtShortSt", id: R.SHORT_INLINE_STRING, value: wide("hi"), cell: "hi" },
	{ name: "BrtShortIsst", id: R.SHORT_STRING, value: u32(1), cell: "second" },
];
const STRING_FORMS = [
	{ name: "BrtCellSt", id: R.INLINE_STRING, header: 8 },
	{ name: "BrtShortSt", id: R.SHORT_INLINE_STRING, header: 4 },
	{ name: "BrtFmlaString", id: R.FORMULA_STRING, header: 8 },
];
const NEXT = short(R.SHORT_BOOL, [0]); // an intact short cell after the cut one: false

describe("parseXlsbSheet — cut cell records (F11.100)", () => {
	it.each(FULL_FORMS)("$name cut at any length below its value has no value", ({ id, value }) => {
		const body = [...u32(3), ...u32(5), ...value];
		for (let length = 0; length < body.length; length++) {
			const sheet = parse(row(2), record(id, body.slice(0, length)), NEXT);
			// A cut column reads as column A and a cut style as style 0, as in 1.0.0.
			const ref = length >= 4 ? "D3" : "A3";
			const next = length >= 4 ? "E3" : "B3";
			expect([...sheet.cells.keys()], `length ${length}`).toEqual([next]);
			expect(sheet.cells.get(next)?.value, `length ${length}`).toBe(false);
			expect(sheet.cellStyles.get(ref), `length ${length}`).toBe(length >= 8 ? 5 : undefined);
		}
	});

	it.each(SHORT_FORMS)(
		"$name cut at any length below its value has no value",
		({ id, value }) => {
			const body = [...u32(5), ...value];
			for (let length = 0; length < body.length; length++) {
				const sheet = parse(
					row(2),
					full(R.BOOL, 2, [1]),
					record(id, body.slice(0, length)),
					NEXT,
				);
				// A cut short record still advances the implicit column (F11.28).
				expect([...sheet.cells.keys()], `length ${length}`).toEqual(["C3", "E3"]);
				expect(sheet.cells.get("E3")?.value, `length ${length}`).toBe(false);
				expect(sheet.cellStyles.get("D3"), `length ${length}`).toBe(
					length >= 4 ? 5 : undefined,
				);
			}
		},
	);

	it.each(FULL_FORMS)(
		"$name reads its value once the value is complete",
		({ id, value, cell }) => {
			const sheet = parse(row(2), full(id, 3, value, 5));
			expect(sheet.cells.get("D3")?.value).toBe(cell);
			expect(sheet.cellStyles.get("D3")).toBe(5);
		},
	);

	it.each(SHORT_FORMS)(
		"$name reads its value once the value is complete",
		({ id, value, cell }) => {
			const sheet = parse(row(2), short(id, value, 5));
			expect(sheet.cells.get("A3")?.value).toBe(cell);
			expect(sheet.cellStyles.get("A3")).toBe(5);
		},
	);

	it("keeps a formula's cached value when only its flags or formula are cut", () => {
		const formula = [0, 0, 1, 2, 3]; // flags and the start of a parsed formula
		expect(parse(full(R.FORMULA_FLOAT, 0, [...real(3.5), ...formula])).cells.get("A1")).toEqual(
			{ ref: "A1", type: "number", value: 3.5 },
		);
		for (const tail of [[], [0], formula]) {
			expect(
				parse(full(R.FORMULA_FLOAT, 0, [...real(3.5), ...tail])).cells.get("A1")?.value,
			).toBe(3.5);
			expect(
				parse(full(R.FORMULA_STRING, 0, [...wide("ok"), ...tail])).cells.get("A1")?.value,
			).toBe("ok");
			expect(parse(full(R.FORMULA_BOOL, 0, [1, ...tail])).cells.get("A1")?.value).toBe(true);
			expect(parse(full(R.FORMULA_BOOLERR, 0, [7, ...tail])).cells.get("A1")?.value).toBe(
				"#DIV/0!",
			);
		}
	});

	it.each(STRING_FORMS)(
		"$name one byte or one code unit short has no value; the null sentinel reads empty",
		({ id, header }) => {
			const make = (payload: readonly number[]) =>
				header === 4 ? short(id, payload, 5) : full(id, 0, payload, 5);
			const value = wide("hi");
			expect(parse(make(value.slice(0, -1))).cells.size).toBe(0); // one byte short
			expect(parse(make(value.slice(0, -2))).cells.size).toBe(0); // one code unit short
			expect(parse(make(value.slice(0, -1))).cellStyles.get("A1")).toBe(5);
			expect(parse(make(value)).cells.get("A1")?.value).toBe("hi");
			expect(parse(make(u32(0xffffffff))).cells.get("A1")?.value).toBe("");
			expect(parse(make(u32(0))).cells.get("A1")?.value).toBe("");
			expect(parse(make(u32(0xffffffff).slice(0, 3))).cells.size).toBe(0);
		},
	);

	it("never replaces an earlier value at the same ref", () => {
		const earlier = parse(
			row(0),
			full(R.NUM, 0, u32((17 << 2) | 2), 2),
			full(R.FLOAT, 0, [1, 2], 6),
		);
		expect(earlier.cells.get("A1")).toEqual({ ref: "A1", type: "number", value: 17 });
		expect(earlier.cellStyles.get("A1")).toBe(6); // the cut record's complete style, as on HEAD
		const short1 = parse(row(0), full(R.BOOL, 0, [1]), row(0), short(R.SHORT_STRING, [], 4));
		expect(short1.cells.get("A1")).toEqual({ ref: "A1", type: "boolean", value: true });
		expect(short1.cellStyles.get("A1")).toBe(4);
		// A cut style is style 0, which clears the earlier style, as on HEAD (out of scope). HEAD
		// read the style byte 3 as the Boolean `true`.
		const cutStyle = parse(row(0), full(R.BOOL, 0, [0], 3), record(R.BOOL, [...u32(0), 3, 0]));
		expect(cutStyle.cells.get("A1")?.value).toBe(false);
		expect(cutStyle.cellStyles.get("A1")).toBeUndefined();
	});

	it("parses cut records in time linear in their count, even a string declaring 0x7FFFFFFF units", () => {
		const sheetOf = (n: number): Uint8Array => {
			const out = [...row(0)];
			for (let i = 0; i < n; i++) {
				out.push(
					...short(R.SHORT_INLINE_STRING, [...u32(0x7fffffff), 65, 0]),
					...short(R.SHORT_STRING, []),
					...full(R.NUM, i % 1000, [1, 2]),
				);
			}
			return Uint8Array.from(out);
		};
		expect(parseXlsbSheet(sheetOf(1000), [], undefined, false).cells.size).toBe(0);
		const ratio = bestDoublingRatio((n) => {
			const bytes = sheetOf(n);
			return () => parseXlsbSheet(bytes, [], undefined, false);
		}, 50_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);
});

describe("valueLength", () => {
	const data = (payload: readonly number[]) => Uint8Array.from(payload);

	it("is 0 for a blank", () => {
		expect(valueLength(R.BLANK, data([]), 8)).toBe(0);
		expect(valueLength(R.SHORT_BLANK, data([]), 4)).toBe(0);
	});

	it("is 4 for an RK number or a shared-string index", () => {
		for (const id of [R.NUM, R.SHORT_NUM, R.STRING, R.SHORT_STRING]) {
			expect(valueLength(id, data([]), 8), String(id)).toBe(4);
		}
	});

	it("is 8 for a real", () => {
		for (const id of [R.FLOAT, R.SHORT_FLOAT, R.FORMULA_FLOAT]) {
			expect(valueLength(id, data([]), 8), String(id)).toBe(8);
		}
	});

	it("is 1 for a Boolean or an error", () => {
		for (const id of [
			R.BOOL,
			R.SHORT_BOOL,
			R.FORMULA_BOOL,
			R.BOOLERR,
			R.SHORT_BOOLERR,
			R.FORMULA_BOOLERR,
		]) {
			expect(valueLength(id, data([]), 8), String(id)).toBe(1);
		}
	});

	it("is the count plus its code units for a wide string, read after the header", () => {
		expect(valueLength(R.SHORT_INLINE_STRING, data([...u32(5), ...u32(3)]), 4)).toBe(10);
		expect(valueLength(R.INLINE_STRING, data([...u32(0), ...u32(5), ...u32(0)]), 8)).toBe(4);
		expect(
			valueLength(R.FORMULA_STRING, data([...u32(0), ...u32(5), ...u32(0x7fffffff)]), 8),
		).toBe(4 + 0x7fffffff * 2);
	});

	it("is 4 for the null sentinel or a cut count", () => {
		expect(valueLength(R.SHORT_INLINE_STRING, data([...u32(5), ...u32(0xffffffff)]), 4)).toBe(
			4,
		);
		expect(valueLength(R.SHORT_INLINE_STRING, data([...u32(5), 2, 0, 0]), 4)).toBe(4);
		expect(valueLength(R.INLINE_STRING, data([]), 8)).toBe(4);
	});
});
