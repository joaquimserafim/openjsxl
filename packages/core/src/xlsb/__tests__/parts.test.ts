import { describe, expect, it } from "vitest";
import { readRecords } from "../../biff";
import { parseXlsbWorkbook } from "../parts";
import { SHEET, WB_PROP } from "../records";

// F11.13 — BrtWbProp.f1904. The workbook part's property record carries the 1904 date system in
// bit 0 of its leading flags word (MS-XLSB §2.4.866). These pins build workbook.bin streams record
// by record: absent, bit clear, bit set, every unrelated bit, truncated payloads, placement, and
// duplicates — the whole tolerant-reader contract for the flag.

// Small BIFF12 builders: byte-combined ids, 7-bit varint lengths, LE integers, XLWideString.
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
const u32b = (v: number): number[] => [
	v & 0xff,
	(v >> 8) & 0xff,
	(v >> 16) & 0xff,
	(v >>> 24) & 0xff,
];
const wstr = (s: string): number[] => {
	const out = u32b(s.length);
	for (let i = 0; i < s.length; i++) {
		const code = s.charCodeAt(i);
		out.push(code & 0xff, code >> 8);
	}
	return out;
};
const rec = (id: number, payload: number[]): number[] => [
	...rid(id),
	...rlen(payload.length),
	...payload,
];

const BEGIN_BOOK = rec(0x0183, []);
const END_BOOK = rec(0x0184, []);
const BEGIN_SHEETS = rec(0x018f, []);
const END_SHEETS = rec(0x0190, []);
const SHEET1 = rec(SHEET, [...u32b(0), ...u32b(1), ...wstr("rId1"), ...wstr("Sheet1")]);
const SHEET_LIST = [...BEGIN_SHEETS, ...SHEET1, ...END_SHEETS];

/** BrtWbProp as Excel lays it out: flags, dwThemeVersion, strName (the workbook code name). */
const wbProp = (flags: number): number[] =>
	rec(WB_PROP, [...u32b(flags), ...u32b(0), ...wstr("ThisWorkbook")]);

/** workbook.bin with the given records between BrtBeginBook and the sheet list. */
const book = (...before: number[][]): Uint8Array =>
	Uint8Array.from([...BEGIN_BOOK, ...before.flat(), ...SHEET_LIST, ...END_BOOK]);

const ONE_SHEET = [{ name: "Sheet1", relId: "rId1", state: "visible" }];

describe("parseXlsbWorkbook — BrtWbProp.f1904 (F11.13)", () => {
	it("frames the record: type 0x0099's varint bytes (99 01) decode to the WB_PROP constant", () => {
		const [first] = [...readRecords(Uint8Array.from([0x99, 0x01, 0x00]))];
		expect(first?.id).toBe(WB_PROP);
		expect(rid(WB_PROP)).toEqual([0x99, 0x01]);
	});

	it("absent → false, and the sheet list is still read", () => {
		expect(parseXlsbWorkbook(book())).toEqual({ sheets: ONE_SHEET, date1904: false });
	});

	it("bit 0 clear → false", () => {
		expect(parseXlsbWorkbook(book(wbProp(0)))).toEqual({ sheets: ONE_SHEET, date1904: false });
	});

	it("bit 0 set → true, and scanning continues past the record to the sheet list", () => {
		expect(parseXlsbWorkbook(book(wbProp(1)))).toEqual({ sheets: ONE_SHEET, date1904: true });
	});

	it("no unrelated or reserved bit enables it; bit 0 alongside them still does", () => {
		// reserved1, fHideBorderUnselLists … fRefreshAll, the reserved high bits: every bit but 0.
		expect(parseXlsbWorkbook(book(wbProp(0xfffffffe))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(wbProp(0x00000002))).date1904).toBe(false); // reserved1
		expect(parseXlsbWorkbook(book(wbProp(0x00000008))).date1904).toBe(false); // fFilterPrivacy
		expect(parseXlsbWorkbook(book(wbProp(0x80000000))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(wbProp(0xffffffff))).date1904).toBe(true);
		expect(parseXlsbWorkbook(book(wbProp(0x00000101))).date1904).toBe(true);
	});

	it("a payload too short for the flags word degrades to false; exactly four bytes suffice", () => {
		expect(parseXlsbWorkbook(book(rec(WB_PROP, []))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(rec(WB_PROP, [0x01]))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(rec(WB_PROP, [0x01, 0x00, 0x00]))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(rec(WB_PROP, [0x01, 0x00, 0x00, 0x00]))).date1904).toBe(true);
		// A declared length that overruns the stream is trimmed by the framer: the flags word is
		// still whole here, so the flag reads; the sheet list that never arrived is simply empty.
		const overrun = Uint8Array.from([
			...BEGIN_BOOK,
			...rid(WB_PROP),
			...rlen(64),
			0x01,
			0,
			0,
			0,
		]);
		expect(parseXlsbWorkbook(overrun)).toEqual({ sheets: [], date1904: true });
	});

	it("is read wherever the record sits, and the last occurrence wins", () => {
		const after = Uint8Array.from([...BEGIN_BOOK, ...SHEET_LIST, ...wbProp(1), ...END_BOOK]);
		expect(parseXlsbWorkbook(after)).toEqual({ sheets: ONE_SHEET, date1904: true });
		expect(parseXlsbWorkbook(book(wbProp(1), wbProp(0))).date1904).toBe(false);
		expect(parseXlsbWorkbook(book(wbProp(0), wbProp(1))).date1904).toBe(true);
	});

	it("stops at BrtEndBook: a property record after it is not the workbook's", () => {
		const trailing = Uint8Array.from([...BEGIN_BOOK, ...SHEET_LIST, ...END_BOOK, ...wbProp(1)]);
		expect(parseXlsbWorkbook(trailing).date1904).toBe(false);
	});
});
