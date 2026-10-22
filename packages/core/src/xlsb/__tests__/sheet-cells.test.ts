import { describe, expect, it } from "vitest";
import { MAX_COL, MAX_ROW } from "../../ooxml";
import type { XlsbStyleTable } from "../parts";
import { parseXlsbSheet } from "../sheet";

// Literal wire IDs keep this producer-compatibility oracle independent of the parser constants.
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
const row = (index: number): number[] => record(0x00, u32(index));
const full = (id: number, col: number, payload: readonly number[], style = 0): number[] =>
	record(id, [...u32(col), ...u32(style), ...payload]);
const short = (id: number, payload: readonly number[], style = 0): number[] =>
	record(id, [...u32(style), ...payload]);
const dimension = (r1 = 0, r2 = 0, c1 = 0, c2 = 0): number[] =>
	record(0x0194, [...u32(r1), ...u32(r2), ...u32(c1), ...u32(c2)]);
const parse = (...records: readonly number[][]) =>
	parseXlsbSheet(Uint8Array.from(records.flat()), ["shared", "第二 🌍"], undefined, false);

const variants = [
	{ name: "blank", fullId: 0x01, shortId: 0x0c, payload: [], type: undefined, value: undefined },
	{
		name: "RK",
		fullId: 0x02,
		shortId: 0x0d,
		payload: u32((17 << 2) | 2),
		type: "number",
		value: 17,
	},
	{ name: "error", fullId: 0x03, shortId: 0x0e, payload: [7], type: "error", value: "#DIV/0!" },
	{ name: "boolean", fullId: 0x04, shortId: 0x0f, payload: [1], type: "boolean", value: true },
	{ name: "real", fullId: 0x05, shortId: 0x10, payload: real(2.75), type: "number", value: 2.75 },
	{
		name: "inline",
		fullId: 0x06,
		shortId: 0x11,
		payload: wide("字 🌍"),
		type: "string",
		value: "字 🌍",
	},
	{ name: "SST", fullId: 0x07, shortId: 0x12, payload: u32(1), type: "string", value: "第二 🌍" },
];

describe("parseXlsbSheet — full and short cell records", () => {
	it.each(variants)("decodes short $name after an explicit cell", ({
		shortId,
		payload,
		type,
		value,
	}) => {
		const sheet = parse(dimension(), row(2), full(0x04, 3, [0]), short(shortId, payload, 7));
		expect(sheet.cellStyles.get("E3")).toBe(7);
		expect(sheet.dimension).toBe("D3:E3");
		expect(sheet.cells.get("E3")).toEqual(
			type === undefined ? undefined : { ref: "E3", type, value },
		);
	});

	it.each(["", "inline", "漢字 🌍", "a".repeat(80)])("decodes full inline string %j", (value) => {
		const sheet = parse(full(0x06, 0, wide(value), 3));
		expect(sheet.cells.get("A1")).toEqual({ ref: "A1", type: "string", value });
		expect(sheet.cellStyles.get("A1")).toBe(3);
	});

	it("starts a short-only row at A and advances through a styled blank", () => {
		const sheet = parse(
			dimension(),
			row(4),
			...variants.map((v) => short(v.shortId, v.payload, 9)),
		);
		expect([...sheet.cells.keys()]).toEqual(["B5", "C5", "D5", "E5", "F5", "G5"]);
		expect([...sheet.cellStyles.keys()]).toEqual(["A5", "B5", "C5", "D5", "E5", "F5", "G5"]);
		expect(sheet.dimension).toBe("A5:G5");
		expect(sheet.cells.get("B5")?.value).toBe(17);
		expect(sheet.cells.get("G5")?.value).toBe("第二 🌍");
	});

	it.each([
		...variants.map(({ name, fullId, payload }) => ({ name, fullId, payload })),
		{ name: "formula string", fullId: 0x08, payload: wide("cached") },
		{ name: "formula number", fullId: 0x09, payload: real(42) },
		{ name: "formula boolean", fullId: 0x0a, payload: [0] },
		{ name: "formula error", fullId: 0x0b, payload: [42] },
	])("sets the next implicit column after a full $name", ({ fullId, payload }) => {
		const sheet = parse(row(0), full(fullId, 5, payload), short(0x0f, [1]));
		expect(sheet.cells.get("G1")).toEqual({ ref: "G1", type: "boolean", value: true });
		expect(sheet.cells.has("A1")).toBe(false);
		expect(sheet.cellStyles.has("F1")).toBe(true);
	});

	it("uses the last explicit column across gaps, backwards records, and blank runs", () => {
		const sheet = parse(
			dimension(),
			full(0x02, 5, u32((1 << 2) | 2)),
			short(0x0c, [], 4),
			short(0x0d, u32((2 << 2) | 2)),
			full(0x02, 1, u32((3 << 2) | 2)),
			short(0x0d, u32((4 << 2) | 2)),
			full(0x01, 10, [], 6),
			short(0x0d, u32((5 << 2) | 2)),
		);
		expect([...sheet.cells].map(([ref, cell]) => [ref, cell.value])).toEqual([
			["F1", 1],
			["H1", 2],
			["B1", 3],
			["C1", 4],
			["L1", 5],
		]);
		expect(sheet.cellStyles.get("G1")).toBe(4);
		expect(sheet.cellStyles.get("K1")).toBe(6);
		expect(sheet.dimension).toBe("B1:L1");
	});

	it("uses an unsupported rich-string cell's explicit column as the next short-cell anchor", () => {
		// SheetJS accepts BrtCellRString in a worksheet: full Cell header, then RichStr flags/text.
		// Value/style decoding for that record stays outside this compatibility fix.
		const sheet = parse(
			dimension(1, 1, 2, 3),
			row(1),
			full(0x3e, 2, [0, ...wide("anchor")], 7),
			short(0x0d, u32((42 << 2) | 2)),
		);
		expect([...sheet.cells]).toEqual([["D2", { ref: "D2", type: "number", value: 42 }]]);
		expect([...sheet.cellStyles]).toEqual([["D2", 0]]);
		expect(sheet.dimension).toBe("C2:D2");
	});

	it("counts rich-string anchors in zero dimensions without exposing their values or styles", () => {
		const anchor = full(0x3e, 2, [0, ...wide("anchor")], 7);
		const alone = parse(dimension(), row(1), anchor);
		expect(alone.dimension).toBe("C2");
		expect(alone.cells.size).toBe(0);
		expect(alone.cellStyles.size).toBe(0);
		const next = parse(dimension(), row(1), anchor, short(0x0f, [1]));
		expect(next.dimension).toBe("C2:D2");
		expect([...next.cells.keys()]).toEqual(["D2"]);
		expect([...next.cellStyles.keys()]).toEqual(["D2"]);
	});

	it("keeps rich-string raw column overflow and row resets consistent with other full cells", () => {
		const sheet = parse(
			dimension(),
			row(0xffffffff),
			full(0x3e, 0xffffffff, [0, ...wide("anchor")]),
			short(0x0d, u32((42 << 2) | 2)),
			row(0),
			short(0x0f, [1]),
		);
		expect(sheet.dimension).toBe("A1:XFD1048576");
		expect([...sheet.cells].map(([ref, cell]) => [ref, cell.value])).toEqual([
			["XFD1048576", 42],
			["A1", true],
		]);
	});

	it("resets the implicit cursor at every row header, including repeated row numbers", () => {
		const sheet = parse(
			short(0x0f, [1]),
			row(2),
			full(0x04, 7, [1]),
			short(0x0f, [0]),
			row(3),
			short(0x0f, [1]),
			short(0x0f, [0]),
			row(3),
			short(0x0f, [0]),
		);
		expect([...sheet.cells.keys()]).toEqual(["A1", "H3", "I3", "A4", "B4"]);
		expect(sheet.cells.get("A4")?.value).toBe(false);
	});

	it("does not let unrelated records alter the cell cursor", () => {
		const sheet = parse(
			full(0x04, 2, [1]),
			record(0x7f, [...u32(99), ...u32(42)]),
			dimension(),
			record(0x003c, [...u32(6), ...u32(8), ...u32(2048), ...u32(0)]),
			record(0x03ee, [...u32(0), ...u32(0), ...u32(0), ...u32(0), ...wide("rId1")]),
			short(0x0f, [0]),
		);
		expect([...sheet.cells.keys()]).toEqual(["C1", "D1"]);
		expect(sheet.dimension).toBe("C1:I1");
		expect(sheet.hyperlinks).toEqual([{ ref: "A1", relId: "rId1" }]);
	});

	it.each([
		{ payload: u32((-17 << 2) | 2), expected: -17 },
		{ payload: u32((-1725 << 2) | 3), expected: -17.25 },
		{ payload: u32(0x400c0000), expected: 3.5 },
		{ payload: u32(0x400c0001), expected: 0.035 },
	])("decodes each RK representation $expected", ({ payload, expected }) => {
		expect(parse(short(0x0d, payload)).cells.get("A1")?.value).toBe(expected);
	});

	it.each([
		false,
		true,
	])("shares date conversion for full and short numeric forms (1904=%s)", (date1904) => {
		const styles: XlsbStyleTable = {
			isDateStyle: (style) => style === 1,
			formatCode: (style) => (style === 1 ? "yyyy-mm-dd" : undefined),
		};
		const bytes = Uint8Array.from([
			...full(0x02, 0, u32((50 << 2) | 3), 0x01000001),
			...short(0x0d, u32((50 << 2) | 3), 0xff000001),
			...full(0x05, 2, real(0.5), 0x80000001),
			...short(0x10, real(0.5), 0x7f000001),
			...full(0x09, 4, real(0.5), 1),
			...short(0x10, real(0.5), 0xff000000),
		]);
		const sheet = parseXlsbSheet(bytes, [], styles, date1904);
		const expected = new Date(
			date1904 ? "1904-01-01T12:00:00.000Z" : "1899-12-30T12:00:00.000Z",
		);
		for (const ref of ["A1", "B1", "C1", "D1", "E1"]) {
			expect(sheet.cells.get(ref)).toEqual({ ref, type: "date", value: expected });
			expect(sheet.cellStyles.get(ref)).toBe(1);
		}
		expect(sheet.cells.get("F1")).toEqual({ ref: "F1", type: "number", value: 0.5 });
		expect(sheet.cellStyles.get("F1")).toBe(0);
	});

	it.each([
		1e308,
		Infinity,
		-Infinity,
		NaN,
	])("preserves a numeric value outside Date range: %s", (value) => {
		const styles: XlsbStyleTable = { isDateStyle: () => true, formatCode: () => "yyyy-mm-dd" };
		const sheet = parseXlsbSheet(Uint8Array.from(short(0x10, real(value))), [], styles, false);
		expect(sheet.cells.get("A1")).toEqual({ ref: "A1", type: "number", value });
	});

	it.each([
		[0, "#NULL!"],
		[7, "#DIV/0!"],
		[15, "#VALUE!"],
		[23, "#REF!"],
		[29, "#NAME?"],
		[36, "#NUM!"],
		[42, "#N/A"],
		[43, "#GETTING_DATA"],
		[255, "#ERR"],
	])("maps short error code %i to %s", (code, value) => {
		if (typeof code !== "number") throw new Error("numeric error code required");
		expect(parse(short(0x0e, [code])).cells.get("A1")).toEqual({
			ref: "A1",
			type: "error",
			value,
		});
	});

	it("preserves empty/Unicode inline strings, SST indexes, and masked style bits", () => {
		const sheet = parse(
			short(0x11, wide(""), 0xffffffff),
			short(0x11, wide("漢字 🌍")),
			short(0x11, u32(0xffffffff)),
			short(0x12, u32(0)),
			short(0x12, u32(1)),
			short(0x12, u32(0xffffffff)),
			short(0x0f, [0]),
		);
		expect([...sheet.cells.values()].map((cell) => cell.value)).toEqual([
			"",
			"漢字 🌍",
			"",
			"shared",
			"第二 🌍",
			"",
			false,
		]);
		expect(sheet.cellStyles.get("A1")).toBe(0xffffff);
	});

	it.each(variants)("keeps full/short $name truncation fallbacks equivalent and advances once", ({
		fullId,
		shortId,
		payload,
	}) => {
		const body = [...u32(0xff000007), ...payload];
		for (let length = 0; length <= body.length; length++) {
			const truncated = body.slice(0, length);
			const explicit = parse(record(fullId, [...u32(0), ...truncated]));
			const implicit = parse(record(shortId, truncated), short(0x0f, [1]));
			expect(implicit.cells.get("A1"), `payload length ${length}`).toEqual(
				explicit.cells.get("A1"),
			);
			expect(implicit.cellStyles.get("A1"), `payload length ${length}`).toBe(
				explicit.cellStyles.get("A1"),
			);
			expect(implicit.cells.get("B1"), `payload length ${length}`).toEqual({
				ref: "B1",
				type: "boolean",
				value: true,
			});
		}
	});

	it("bounds declared inline lengths to the actual record and keeps following records intact", () => {
		const sheet = parse(short(0x11, [...u32(0xfffffffe), 65, 0, 66]), short(0x0f, [1]));
		expect(sheet.cells.get("A1")?.value).toBe("A�");
		expect(sheet.cells.get("B1")?.value).toBe(true);
		const overrun = parseXlsbSheet(
			Uint8Array.from([0x11, 127, ...u32(0), ...wide("end")]),
			[],
			undefined,
			false,
		);
		expect(overrun.cells.get("A1")?.value).toBe("end");
	});

	it("does not wrap raw columns at uint32 overflow and clamps only cell references", () => {
		const sheet = parse(
			dimension(),
			row(0xffffffff),
			full(0x04, 0xffffffff, [1]),
			short(0x0c, [], 3),
			short(0x0f, [0], 9),
			row(MAX_ROW - 1),
			short(0x0f, [1]),
		);
		expect([...sheet.cells.keys()]).toEqual(["XFD1048576", "A1048576"]);
		expect(sheet.cells.get("XFD1048576")?.value).toBe(false);
		expect(sheet.cellStyles.get("XFD1048576")).toBe(9);
		expect(sheet.dimension).toBe("A1048576:XFD1048576");
		const edge = parse(full(0x04, MAX_COL - 2, [1]), short(0x0f, [1]), short(0x0f, [0]));
		expect([...edge.cells.keys()]).toEqual(["XFC1", "XFD1"]);
		expect(edge.cells.get("XFD1")?.value).toBe(false);
	});

	it.each([
		{ name: "absent", prefix: [], expected: undefined },
		{ name: "zero", prefix: dimension(), expected: "A1:D4" },
		{
			name: "all-max",
			prefix: dimension(0xffffffff, 0xffffffff, 0xffffffff, 0xffffffff),
			expected: undefined,
		},
		{ name: "ordinary", prefix: dimension(2, 5, 1, 2), expected: "B3:C6" },
	])("preserves the $name dimension policy with short cells", ({ prefix, expected }) => {
		const sheet = parse(
			prefix,
			short(0x0c, [], 0),
			row(3),
			full(0x04, 2, [1]),
			short(0x0f, [0]),
		);
		expect(sheet.dimension).toBe(expected);
		expect(sheet.cellStyles.get("A1")).toBe(0);
		expect(sheet.cells.get("D4")?.value).toBe(false);
	});

	it("resolves zero dimensions after scanning and retains last-dimension-wins", () => {
		const cells = [row(4), short(0x0c, [], 0), short(0x0f, [1])];
		expect(parse(...cells, dimension()).dimension).toBe("A5:B5");
		expect(parse(dimension(), ...cells, dimension(2, 2, 2, 2)).dimension).toBe("C3");
		expect(parse(dimension(2, 2, 2, 2), ...cells, dimension()).dimension).toBe("A5:B5");
	});
});
