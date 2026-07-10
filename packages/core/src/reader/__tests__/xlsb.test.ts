import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { RecordData, readRecords } from "../../biff";
import { XlsxError } from "../../errors";
import { workbookToInput, writeXlsx } from "../../writer";
import { parseXlsbSheet, parseXlsbStyles } from "../../xlsb";
import { openXlsx } from "../workbook";
import { openXlsb } from "../xlsb";

// F7.2 — the .xlsb reader. Values are cross-checked against python-calamine + pyxlsb (both read the
// committed xlsb-basic.xlsb; see data/README.md provenance). The BIFF12 record layer is unit-tested
// directly for the RK corners and hostile framing.

describe("openXlsb — xlsb-basic.xlsb (values cross-checked vs calamine/pyxlsb)", () => {
	it("lists sheets in tab order, with visibility", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-basic.xlsb"));
		expect(wb.sheets.map((s) => [s.name, s.state, s.visible])).toEqual([
			["Sheet1", "visible", true],
			["Sheet2", "visible", true],
			["Hidden", "hidden", false],
		]);
	});

	it("reads every cell type", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		expect(s.cell("A1")).toEqual({ ref: "A1", type: "string", value: "hello" });
		expect(s.cell("B1")).toEqual({ ref: "B1", type: "number", value: 42 }); // RK int
		expect(s.cell("C1")).toEqual({ ref: "C1", type: "number", value: 3.14159 }); // real double
		expect(s.cell("D1")).toEqual({ ref: "D1", type: "number", value: 0.25 }); // RK ÷100
		expect(s.cell("F1")).toEqual({ ref: "F1", type: "boolean", value: true });
		expect(s.cell("G1")).toEqual({ ref: "G1", type: "error", value: "#DIV/0!" });
	});

	it("detects a date-styled number via the style table (builtin numFmt 14)", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		const e1 = s.cell("E1");
		expect(e1.type).toBe("date");
		expect((e1.value as Date).getTime()).toBe(Date.UTC(2020, 0, 1)); // serial 43831
		expect(s.numberFormat("E1")).toBe("mm-dd-yy"); // builtin 14
	});

	it("reads cached formula results (number and string)", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		expect(s.cell("A2")).toEqual({ ref: "A2", type: "number", value: 84 });
		expect(s.cell("B2")).toEqual({ ref: "B2", type: "string", value: "cached" });
		expect(s.formula("A2")).toBeUndefined(); // formula TEXT is not carried
	});

	it("reads a hyperlink resolved through the sheet rels, and the dimension", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		expect(s.hyperlinks).toEqual([{ ref: "H1", target: "https://example.com/x" }]);
		expect(s.dimension).toBe("A1:G2");
	});

	it("reads a second sheet and a hidden sheet", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-basic.xlsb"));
		expect(wb.sheet("Sheet2").cell("A1").value).toBe("second");
		expect(wb.sheet("Hidden").cell("A1").value).toBe("hello");
	});

	it("streams rows in ascending row/column order", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		const rows = [];
		for await (const row of s.rows()) rows.push([row.index, row.cells.map((c) => c.ref)]);
		expect(rows[0]).toEqual([1, ["A1", "B1", "C1", "D1", "E1", "F1", "G1"]]);
		expect(rows[1]).toEqual([2, ["A2", "B2"]]);
	});

	it("degrades unsupported accessors (styles/comments/geometry/images/merges)", async () => {
		const s = (await openXlsb(await loadFixture("xlsb-basic.xlsb"))).sheet("Sheet1");
		expect(s.style("A1")).toBeUndefined();
		expect(s.mergedCells).toEqual([]);
		expect(s.comments).toEqual([]);
		expect(s.columns).toEqual([]);
		expect(s.rowProperties.size).toBe(0);
		expect(s.freeze).toBeUndefined();
		expect(await s.images()).toEqual([]);
	});

	it("converts .xlsb to .xlsx through the bridge, values intact", async () => {
		const wb = await openXlsb(await loadFixture("xlsb-basic.xlsb"));
		const out = await openXlsx(await writeXlsx(await workbookToInput(wb)));
		const s = out.sheet("Sheet1");
		expect(s.cell("A1").value).toBe("hello");
		expect(s.cell("B1").value).toBe(42);
		expect(s.cell("E1").type).toBe("date");
		expect(out.sheets.map((x) => x.name)).toEqual(["Sheet1", "Sheet2", "Hidden"]);
	});
});

describe("openXlsb — typed failures", () => {
	it("throws not-a-zip on non-ZIP bytes", async () => {
		await expect(openXlsb(new Uint8Array([1, 2, 3, 4]))).rejects.toMatchObject({
			name: "XlsxError",
			code: "not-a-zip",
		});
	});

	it("surfaces failures as XlsxError", async () => {
		const err = await openXlsb(new Uint8Array([0])).catch((e) => e);
		expect(err).toBeInstanceOf(XlsxError);
	});
});

// A little-endian i32 as 4 bytes, for building RK payloads.
const i32 = (n: number): Uint8Array => {
	const b = new Uint8Array(4);
	new DataView(b.buffer).setInt32(0, n, true);
	return b;
};

describe("biff record layer — RK decode + framing", () => {
	it("decodes RK integers, ÷100, negatives, and truncated doubles", () => {
		expect(new RecordData(i32((42 << 2) | 0x02)).rk()).toBe(42); // int
		expect(new RecordData(i32((25 << 2) | 0x03)).rk()).toBe(0.25); // int ÷100
		expect(new RecordData(i32((-5 << 2) | 0x02)).rk()).toBe(-5); // negative int (arithmetic shift)
		// 1.5 = 0x3FF8000000000000; its high 4 bytes are the RK payload (low 2 bits already 0).
		expect(new RecordData(i32(0x3ff80000)).rk()).toBe(1.5);
	});

	it("reads uint fields and a UTF-16LE wide string", () => {
		const bytes = new Uint8Array([
			0x2a,
			0x00,
			0x00,
			0x00, // u32 = 42
			0x02,
			0x00,
			0x00,
			0x00, // wide-string count = 2
			0x68,
			0x00,
			0x69,
			0x00, // "hi" UTF-16LE
		]);
		const d = new RecordData(bytes);
		expect(d.u32()).toBe(42);
		expect(d.wideString()).toBe("hi");
	});

	it("returns the null-string sentinel as undefined", () => {
		expect(new RecordData(i32(-1)).wideString()).toBeUndefined(); // 0xFFFFFFFF
	});

	it("degrades a short record: reads past the end give 0 / empty, never throw", () => {
		const d = new RecordData(new Uint8Array([1, 2])); // only 2 bytes
		expect(d.u32()).toBe(0); // not enough bytes → 0
		expect(d.wideString()).toBeUndefined();
	});

	it("frames records and trims a lying/oversized length to the buffer (no OOB)", () => {
		// record A: id 0x0002, len 1, payload [0xAA]; record B: id 0x0002, len 100 (lie), payload [0x01]
		const stream = new Uint8Array([0x02, 0x01, 0xaa, 0x02, 0x64, 0x01]);
		const recs = [...readRecords(stream)];
		expect(recs).toHaveLength(2);
		expect(recs[0]).toMatchObject({ id: 0x02 });
		expect([...(recs[0]?.data ?? [])]).toEqual([0xaa]);
		// second record's declared length (100) overran the buffer → trimmed to what remained (1 byte)
		expect([...(recs[1]?.data ?? [])]).toEqual([0x01]);
	});

	it("decodes a 2-byte record id (byte-combined, like Excel writes)", () => {
		// id 0x03EE (BrtHLink) writes as [0xEE, 0x03]; len 0
		const recs = [...readRecords(new Uint8Array([0xee, 0x03, 0x00]))];
		expect(recs[0]?.id).toBe(0x03ee);
	});
});

// Small BIFF12 record builders for the regression below.
const rid = (c: number): number[] => (c < 0x80 ? [c] : [c & 0xff, c >> 8]);
const rlen = (n: number): number[] => {
	const out: number[] = [];
	let v = n;
	do {
		let b = v & 0x7f;
		v >>>= 7;
		if (v) b |= 0x80;
		out.push(b);
	} while (v);
	return out;
};
const u16b = (v: number): number[] => [v & 0xff, (v >> 8) & 0xff];
const u32b = (v: number): number[] => [
	v & 0xff,
	(v >> 8) & 0xff,
	(v >> 16) & 0xff,
	(v >>> 24) & 0xff,
];
const rec = (id: number, payload: number[]): number[] => [
	...rid(id),
	...rlen(payload.length),
	...payload,
];

describe("xlsb — cell style field is masked to 24-bit iStyleRef (adversarial-review regression)", () => {
	it("resolves a date format even when fPhShow (bit 24) is set on the style field", () => {
		// styles.bin: cellXfs [xf0→numFmt 0, xf1→builtin numFmt 14 (a date)].
		const styles = parseXlsbStyles(
			Uint8Array.from([
				...rec(0x0296, []), // STYLESHEET
				...rec(0x04e9, u32b(2)), // CELLXFS count=2
				...rec(0x002f, [
					...u16b(0xffff),
					...u16b(0),
					...u16b(0),
					...u16b(0),
					...u16b(0),
					0,
					0,
					...u16b(0),
				]),
				...rec(0x002f, [
					...u16b(0xffff),
					...u16b(14),
					...u16b(0),
					...u16b(0),
					...u16b(0),
					0,
					0,
					...u16b(0),
				]),
				...rec(0x04ea, []), // CELLXFS_END
				...rec(0x0297, []), // STYLESHEET_END
			]),
		);
		// sheet.bin: A1 = an RK date serial (43831) whose style field is 0x01000001 —
		// fPhShow=1 plus iStyleRef=1. Reading the full u32 would give 16777217 and lose the date.
		const rk = u32b(((43831 << 2) | 0x02) >>> 0);
		const sheet = Uint8Array.from([
			...rec(0x0000, [...u32b(0), ...new Array(12).fill(0)]), // ROW rw=0
			...rec(0x0002, [...u32b(0), ...u32b(0x01000001), ...rk]), // NUM col 0, packed style, date serial
		]);
		const a1 = parseXlsbSheet(sheet, [], styles, false).cells.get("A1");
		expect(a1?.type).toBe("date");
		expect((a1?.value as Date).getTime()).toBe(Date.UTC(2020, 0, 1));
	});
});
