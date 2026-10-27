import { describe, expect, it } from "vitest";
import { MAX_COL, MAX_ROW } from "../../ooxml";
import * as R from "../records";
import { parseXlsbSheet } from "../sheet";

// Synthetic BIFF12 records: byte-combined record ids, varint lengths, and little-endian fields.
const u32 = (n: number): number[] => [
	n & 0xff,
	(n >>> 8) & 0xff,
	(n >>> 16) & 0xff,
	(n >>> 24) & 0xff,
];
function record(id: number, data: readonly number[]): number[] {
	const length: number[] = [];
	let remaining = data.length;
	do {
		const byte = remaining & 0x7f;
		remaining >>>= 7;
		length.push(byte | (remaining === 0 ? 0 : 0x80));
	} while (remaining !== 0);
	return [...(id < 0x80 ? [id] : [id & 0xff, id >>> 8]), ...length, ...data];
}
function real(value: number): number[] {
	const bytes = new Uint8Array(8);
	new DataView(bytes.buffer).setFloat64(0, value, true);
	return [...bytes];
}
function wide(value: string): number[] {
	const bytes = u32(value.length);
	for (let i = 0; i < value.length; i++) {
		const code = value.charCodeAt(i);
		bytes.push(code & 0xff, code >>> 8);
	}
	return bytes;
}
const dimension = (r1 = 0, r2 = 0, c1 = 0, c2 = 0): number[] =>
	record(R.DIMENSION, [...u32(r1), ...u32(r2), ...u32(c1), ...u32(c2)]);
const row = (index: number): number[] => record(R.ROW, u32(index));
const formattedRow = (index: number, style = 1, flags = 0x40): number[] =>
	record(R.ROW, [...u32(index), ...u32(style), 0x2c, 0x01, 0, flags, 0, ...u32(0)]);
const formattedColumns = (first: number, last: number, style = 1): number[] =>
	record(0x003c, [...u32(first), ...u32(last), ...u32(2048), ...u32(style), 0, 0]);
const cell = (id: number, col = 0, style = 0, payload: readonly number[] = []): number[] =>
	record(id, [...u32(col), ...u32(style), ...payload]);
const parse = (...records: readonly number[][]) =>
	parseXlsbSheet(Uint8Array.from(records.flat()), ["shared"], undefined, false);
const INVALID_DIMENSION = dimension(0xffffffff, 0xffffffff, 0xffffffff, 0xffffffff);

describe("parseXlsbSheet — used dimensions", () => {
	it("treats an all-zero dimension without cell records as empty", () => {
		const sheet = parse(dimension());

		expect(sheet.dimension).toBeUndefined();
		expect(sheet.cells.size).toBe(0);
		expect(sheet.cellStyles.size).toBe(0);
	});

	it("does not treat row records, hyperlinks, or unknown records as used cells", () => {
		const sheet = parse(
			dimension(),
			row(5),
			record(R.HYPERLINK, [...u32(2), ...u32(3), ...u32(2), ...u32(2), ...wide("rId1")]),
			record(0x007f, [...u32(5), ...u32(1)]),
			record(R.SHEETDATA, []),
			record(R.SHEETDATA_END, []),
		);

		expect(sheet.dimension).toBeUndefined();
		expect(sheet.cells.size).toBe(0);
		expect(sheet.cellStyles.size).toBe(0);
		expect(sheet.hyperlinks).toEqual([{ ref: "C3:C4", relId: "rId1" }]);
	});

	it("retains A1 when a numeric cell contradicts the zero dimension", () => {
		const sheet = parse(dimension(), row(0), cell(R.NUM, 0, 0, u32((42 << 2) | 2)));

		expect(sheet.dimension).toBe("A1");
		expect(sheet.cells.get("A1")).toEqual({ ref: "A1", type: "number", value: 42 });
	});

	it.each(["", "inline"])("counts and decodes an inline string %j", (value) => {
		const sheet = parse(dimension(), row(0), cell(0x0006, 0, 7, wide(value)));

		expect(sheet.dimension).toBe("A1");
		expect([...sheet.cells.values()]).toEqual([{ ref: "A1", type: "string", value }]);
		expect([...sheet.cellStyles]).toEqual([["A1", 7]]);
	});

	it("merges sparse inline strings with other cells and retains their values and styles", () => {
		const sheet = parse(
			dimension(),
			row(7),
			cell(0x0006, 4, 7, wide("inline")),
			row(2),
			cell(R.BOOL, 1, 0, [1]),
		);

		expect(sheet.dimension).toBe("B3:E8");
		expect([...sheet.cells]).toEqual([
			["E8", { ref: "E8", type: "string", value: "inline" }],
			["B3", { ref: "B3", type: "boolean", value: true }],
		]);
		expect([...sheet.cellStyles]).toEqual([["E8", 7]]); // B3's style 0 is not stored (F11.59)
	});

	it("keeps absent, invalid, and ordinary dimensions authoritative with inline strings", () => {
		const inline = [row(7), cell(0x0006, 4, 7, wide("inline"))];

		expect(parse(...inline).dimension).toBeUndefined();
		expect(parse(INVALID_DIMENSION, ...inline).dimension).toBeUndefined();
		expect(parse(dimension(3, 6, 1, 1), ...inline).dimension).toBe("B4:B7");
	});

	it.each([0, 7, 0x01000001])(
		"counts an A1 blank with packed style %i as used without inventing a value",
		(style) => {
			const sheet = parse(dimension(), row(0), cell(R.BLANK, 0, style));

			expect(sheet.dimension).toBe("A1");
			expect(sheet.cells.size).toBe(0);
			expect(sheet.cellStyles.get("A1") ?? 0).toBe(style & 0xffffff); // 0 is not stored (F11.59)
		},
	);

	it.each([
		{ name: "RK number", id: R.NUM, payload: u32((7 << 2) | 2), type: "number", value: 7 },
		{ name: "real number", id: R.FLOAT, payload: real(2.5), type: "number", value: 2.5 },
		{ name: "shared string", id: R.STRING, payload: u32(0), type: "string", value: "shared" },
		{ name: "boolean", id: R.BOOL, payload: [1], type: "boolean", value: true },
		{ name: "error", id: R.BOOLERR, payload: [7], type: "error", value: "#DIV/0!" },
		{
			name: "formula string",
			id: R.FORMULA_STRING,
			payload: wide("cached"),
			type: "string",
			value: "cached",
		},
		{
			name: "formula number",
			id: R.FORMULA_FLOAT,
			payload: real(0),
			type: "number",
			value: 0,
		},
		{
			name: "formula boolean",
			id: R.FORMULA_BOOL,
			payload: [0],
			type: "boolean",
			value: false,
		},
		{
			name: "formula error",
			id: R.FORMULA_BOOLERR,
			payload: [42],
			type: "error",
			value: "#N/A",
		},
	])("derives a non-A1 dimension from a $name record", ({ id, payload, type, value }) => {
		const sheet = parse(dimension(), row(6), cell(id, 4, 0, payload));

		expect(sheet.dimension).toBe("E7");
		expect(sheet.cells.get("E7")).toEqual({ ref: "E7", type, value });
		expect(sheet.cellStyles.get("E7") ?? 0).toBe(0); // style 0 is not stored (F11.59)
	});

	it("derives sparse bounds across unsorted rows, values, and styled blanks", () => {
		const sheet = parse(
			dimension(),
			row(8),
			cell(R.BOOL, 4, 0, [1]),
			row(3),
			cell(R.BLANK, 1, 0),
			row(5),
			cell(R.STRING, 3, 0, u32(0)),
		);

		expect(sheet.dimension).toBe("B4:E9");
		expect([...sheet.cells.keys()]).toEqual(["E9", "D6"]);
		expect([...sheet.cellStyles.keys()]).toEqual([]); // every style here is 0, not stored (F11.59)
	});

	it("clamps observed cell coordinates with the same grid bounds as cell placement", () => {
		const sheet = parse(
			dimension(),
			row(2),
			cell(R.BLANK, 1),
			row(0xffffffff),
			cell(R.BOOL, 0xffffffff, 0, [1]),
		);

		expect(sheet.dimension).toBe("B3:XFD1048576");
		expect(sheet.cells.get("XFD1048576")?.value).toBe(true);
	});

	it("uses row zero before the first ROW and retains coordinates after duplicate cells", () => {
		const sheet = parse(
			dimension(),
			cell(R.BOOL, 2, 0, [1]),
			cell(R.BLANK, 2, 9),
			row(4),
			cell(R.BOOL, 1, 0, [0]),
		);

		expect(sheet.dimension).toBe("B1:C5");
		expect(sheet.cells.get("C1")?.value).toBe(true);
		expect(sheet.cellStyles.get("C1")).toBe(9);
	});

	it("keeps an absent dimension absent even when cells are present", () => {
		expect(parse().dimension).toBeUndefined();
		const sheet = parse(row(3), cell(R.BOOL, 1, 0, [1]));

		expect(sheet.dimension).toBeUndefined();
		expect(sheet.cells.get("B4")?.value).toBe(true);
	});

	it.each([false, true])(
		"keeps the invalid all-max sentinel absent (cells present: %s)",
		(withCells) => {
			const records = [INVALID_DIMENSION];
			if (withCells) records.push(row(3), cell(R.BOOL, 1, 0, [1]));
			const sheet = parse(...records);

			expect(sheet.dimension).toBeUndefined();
			expect(sheet.cells.size).toBe(withCells ? 1 : 0);
		},
	);

	it.each([
		{ name: "ordinary", fields: [3, 6, 1, 1], expected: "B4:B7" },
		{ name: "nonzero singleton", fields: [3, 3, 1, 1], expected: "B4" },
		{
			name: "genuine last grid cell",
			fields: [MAX_ROW - 1, MAX_ROW - 1, MAX_COL - 1, MAX_COL - 1],
			expected: "XFD1048576",
		},
		{
			name: "clipped end",
			fields: [3, MAX_ROW + 50, 1, MAX_COL + 50],
			expected: "B4:XFD1048576",
		},
		{
			name: "both ends clipped",
			fields: [MAX_ROW + 1, MAX_ROW + 2, MAX_COL + 1, MAX_COL + 2],
			expected: "XFD1048576",
		},
		{ name: "reversed", fields: [6, 3, 4, 1], expected: "E7:B4" },
		{
			name: "mixed max",
			fields: [0, 0xffffffff, 0xffffffff, 0],
			expected: "XFD1:A1048576",
		},
	])("preserves $name explicit dimension interpretation", ({ fields, expected }) => {
		const sheet = parse(
			record(R.DIMENSION, fields.flatMap(u32)),
			row(0),
			cell(R.BOOL, 0, 0, [1]),
		);

		expect(sheet.dimension).toBe(expected);
		expect(sheet.cells.get("A1")?.value).toBe(true);
	});

	it("reconciles zero dimensions regardless of their position among cells", () => {
		const used = [row(3), cell(R.BOOL, 1, 0, [1]), row(6), cell(R.BLANK, 4)];
		expect(parse(dimension(), ...used).dimension).toBe("B4:E7");
		expect(parse(...used, dimension()).dimension).toBe("B4:E7");
		expect(parse(...used.slice(0, 2), dimension(), ...used.slice(2)).dimension).toBe("B4:E7");
	});

	it.each([
		{
			name: "zero then ordinary",
			dims: [dimension(), dimension(3, 6, 1, 1)],
			expected: "B4:B7",
		},
		{ name: "ordinary then zero", dims: [dimension(3, 6, 1, 1), dimension()], expected: "C9" },
		{ name: "zero then invalid", dims: [dimension(), INVALID_DIMENSION], expected: undefined },
		{ name: "invalid then zero", dims: [INVALID_DIMENSION, dimension()], expected: "C9" },
		{
			name: "ordinary then invalid",
			dims: [dimension(3, 6, 1, 1), INVALID_DIMENSION],
			expected: undefined,
		},
		{
			name: "invalid then ordinary",
			dims: [INVALID_DIMENSION, dimension(3, 6, 1, 1)],
			expected: "B4:B7",
		},
	])("retains last dimension record policy for $name", ({ dims, expected }) => {
		const sheet = parse(row(8), cell(R.BOOL, 2, 0, [1]), ...dims);

		expect(sheet.dimension).toBe(expected);
		expect(sheet.cells.get("C9")?.value).toBe(true);
	});

	it("clears an earlier ordinary dimension when the final zero record has no used cells", () => {
		expect(parse(dimension(3, 6, 1, 1), dimension()).dimension).toBeUndefined();
	});

	it("keeps short DIMENSION payloads tolerant using zero for missing fields", () => {
		for (const payload of [[], [0], [0, 0, 0]]) {
			expect(parse(record(R.DIMENSION, payload)).dimension).toBeUndefined();
			expect(parse(row(3), cell(R.BLANK, 1), record(R.DIMENSION, payload)).dimension).toBe(
				"B4",
			);
		}
		expect(parse(record(R.DIMENSION, [...u32(3), ...u32(6)])).dimension).toBe("A4:A7");
		expect(
			parse(record(R.DIMENSION, [...u32(0xffffffff), ...u32(0xffffffff), ...u32(0xffffffff)]))
				.dimension,
		).toBe("XFD1048576:A1048576");
	});

	it("recognizes truncated cell payloads using existing zero-default decoding", () => {
		const numeric = parse(dimension(), row(5), record(R.NUM, []));
		expect(numeric.dimension).toBe("A6");
		expect(numeric.cells.get("A6")).toEqual({ ref: "A6", type: "number", value: 0 });
		const blank = parse(dimension(), row(3), record(R.BLANK, [1, 0, 0]));
		expect(blank.dimension).toBe("A4");
		expect(blank.cells.size).toBe(0);
		expect(blank.cellStyles.get("A4") ?? 0).toBe(0); // style 0 is not stored (F11.59)
	});

	it("trims a lying final record length and ignores incomplete record headers", () => {
		const used = [row(3), cell(R.BOOL, 1, 0, [1])];
		const overrun = [0x94, 0x01, 64, ...u32(0)];
		expect(parse(...used, overrun).dimension).toBe("B4");
		expect(parse(dimension(3, 6, 1, 1), [0x94]).dimension).toBe("B4:B7");
		expect(parse(dimension(3, 6, 1, 1), [0x94, 0x01, 0x80]).dimension).toBe("B4:B7");
	});

	it("retains A1 for an applied row format without individual cell records", () => {
		const sheet = parse(dimension(), formattedRow(0));

		expect(sheet.dimension).toBe("A1");
		expect(sheet.cells.size).toBe(0);
		expect(sheet.cellStyles.size).toBe(0);
	});

	it("retains A1 for an explicit column format without individual cell records", () => {
		const sheet = parse(dimension(), formattedColumns(0, 0));

		expect(sheet.dimension).toBe("A1");
		expect(sheet.cells.size).toBe(0);
		expect(sheet.cellStyles.size).toBe(0);
	});

	it.each([0, 1])("counts row and column default style %i at non-A1 coordinates", (style) => {
		expect(parse(dimension(), formattedRow(4, style)).dimension).toBe("A5");
		expect(parse(dimension(), formattedColumns(2, 3, style)).dimension).toBe("C1:D1");
	});

	it("combines formatted axes without expanding either to A1", () => {
		expect(parse(dimension(), formattedRow(4), formattedColumns(3, 3)).dimension).toBe("D5");
		expect(parse(dimension(), formattedColumns(3, 3), formattedRow(4)).dimension).toBe("D5");
		expect(parse(formattedColumns(3, 3), formattedRow(4), dimension()).dimension).toBe("D5");
		expect(parse(dimension(), formattedRow(4), formattedColumns(2, 3)).dimension).toBe("C5:D5");
	});

	it("combines default formatting and cell bounds independently on each axis", () => {
		const atC2 = [row(1), cell(R.BOOL, 2, 0, [1])];
		const atB5 = [row(4), cell(R.BOOL, 1, 0, [1])];
		expect(parse(dimension(), formattedRow(4), ...atC2).dimension).toBe("C2:C5");
		expect(parse(dimension(), ...atC2, formattedRow(4)).dimension).toBe("C2:C5");
		expect(parse(dimension(), formattedColumns(3, 3), ...atB5).dimension).toBe("B5:D5");
		expect(parse(dimension(), ...atB5, formattedColumns(3, 3)).dimension).toBe("B5:D5");
		expect(parse(dimension(), formattedRow(4), formattedColumns(3, 3), ...atC2).dimension).toBe(
			"C2:D5",
		);
	});

	it("merges several formatted rows and columns without materializing cells", () => {
		const sheet = parse(
			dimension(),
			formattedRow(16),
			formattedRow(4),
			formattedColumns(6, 7),
			formattedColumns(2, 3),
		);

		expect(sheet.dimension).toBe("C5:H17");
		expect(sheet.cells.size).toBe(0);
		expect(sheet.cellStyles.size).toBe(0);
	});

	it.each([0, 1, 8, 0x10, 0x20, 0x80])(
		"does not count row flags %i without the applied-format bit",
		(flags) => {
			expect(parse(dimension(), formattedRow(4, 1, flags)).dimension).toBeUndefined();
		},
	);

	it.each([0x40, 0x41, 0x50, 0x60, 0xc0])(
		"counts the applied row format alongside unrelated flags %i",
		(flags) => {
			expect(parse(dimension(), formattedRow(4, 1, flags)).dimension).toBe("A5");
		},
	);

	it("reads the applied row-format flag only from its complete byte", () => {
		const payload = [...u32(4), ...u32(1), 0x2c, 0x01, 0, 0x40, 0, ...u32(0)];
		for (const length of [0, 3, 4, 8, 10, 11]) {
			expect(
				parse(dimension(), record(R.ROW, payload.slice(0, length))).dimension,
			).toBeUndefined();
		}
		expect(parse(dimension(), record(R.ROW, payload.slice(0, 12))).dimension).toBe("A5");
		const wrongBytes = [...u32(4), ...u32(1), 0x2c, 0x01, 0x40, 0, 0x40, ...u32(0)];
		expect(parse(dimension(), record(R.ROW, wrongBytes)).dimension).toBeUndefined();
	});

	it("requires a complete column style field but tolerates absent trailing flags", () => {
		const payload = [...u32(2), ...u32(3), ...u32(2048), ...u32(0), 0, 0];
		for (const length of [0, 4, 8, 12, 15]) {
			expect(
				parse(dimension(), record(0x003c, payload.slice(0, length))).dimension,
			).toBeUndefined();
		}
		expect(parse(dimension(), record(0x003c, payload.slice(0, 16))).dimension).toBe("C1:D1");
	});

	it("ignores reversed column defaults even when both endpoints would clamp together", () => {
		expect(parse(dimension(), formattedColumns(3, 2)).dimension).toBeUndefined();
		expect(
			parse(dimension(), formattedColumns(MAX_COL + 5, MAX_COL + 3)).dimension,
		).toBeUndefined();
		expect(parse(dimension(), formattedRow(4), formattedColumns(3, 2)).dimension).toBe("A5");
	});

	it("clamps formatted axes without expanding their coordinate ranges", () => {
		expect(parse(dimension(), formattedRow(0xffffffff)).dimension).toBe("A1048576");
		expect(parse(dimension(), formattedColumns(MAX_COL + 1, 0xffffffff)).dimension).toBe(
			"XFD1",
		);
		expect(
			parse(dimension(), formattedRow(0xffffffff), formattedColumns(0xffffffff, 0xffffffff))
				.dimension,
		).toBe("XFD1048576");
	});

	it("retains absent, invalid, and authoritative dimensions with default formatting", () => {
		const defaults = [formattedRow(4), formattedColumns(3, 3)];
		expect(parse(...defaults).dimension).toBeUndefined();
		expect(parse(INVALID_DIMENSION, ...defaults).dimension).toBeUndefined();
		expect(parse(...defaults, INVALID_DIMENSION).dimension).toBeUndefined();
		expect(parse(dimension(3, 6, 1, 1), ...defaults).dimension).toBe("B4:B7");
		expect(parse(dimension(), ...defaults, dimension(3, 6, 1, 1)).dimension).toBe("B4:B7");
		expect(parse(dimension(3, 6, 1, 1), ...defaults, dimension()).dimension).toBe("D5");
		expect(parse(dimension(), ...defaults, INVALID_DIMENSION).dimension).toBeUndefined();
		expect(parse(INVALID_DIMENSION, ...defaults, dimension()).dimension).toBe("D5");
	});
});
