import { describe, expect, it } from "vitest";
import { dateToSerial, serialToDate } from "../dates";

describe("serialToDate", () => {
	it("maps the 1900-system serial for the Unix epoch", () => {
		expect(serialToDate(25569).getTime()).toBe(Date.UTC(1970, 0, 1));
	});

	it("maps a modern 1900-system date", () => {
		expect(serialToDate(43831).getTime()).toBe(Date.UTC(2020, 0, 1));
	});

	it("honours the 1904 date system", () => {
		expect(serialToDate(42369, true).getTime()).toBe(Date.UTC(2020, 0, 1));
	});

	it("decodes fractional serials as times", () => {
		expect(serialToDate(43831.5).getTime()).toBe(Date.UTC(2020, 0, 1, 12));
	});
});

describe("dateToSerial", () => {
	it("is the inverse of serialToDate for the 1900 system", () => {
		expect(dateToSerial(new Date(Date.UTC(2020, 0, 1)))).toBe(43831);
		expect(dateToSerial(new Date(Date.UTC(1970, 0, 1)))).toBe(25569);
	});

	it("honours the 1904 date system", () => {
		expect(dateToSerial(new Date(Date.UTC(2020, 0, 1)), true)).toBe(42369);
	});

	it("encodes a time-of-day as a fractional serial", () => {
		expect(dateToSerial(new Date(Date.UTC(2020, 0, 1, 12)))).toBe(43831.5);
	});

	it("round-trips a range of dates through serialToDate losslessly", () => {
		const samples = [
			Date.UTC(1900, 2, 1),
			Date.UTC(1970, 0, 1),
			Date.UTC(1999, 11, 31, 23, 59, 59),
			Date.UTC(2020, 0, 1, 6, 30),
			Date.UTC(2026, 6, 1),
			Date.UTC(2099, 11, 31),
		];
		for (const ms of samples) {
			const d = new Date(ms);
			expect(serialToDate(dateToSerial(d)).getTime()).toBe(ms);
			expect(serialToDate(dateToSerial(d, true), true).getTime()).toBe(ms);
		}
	});

	it("returns NaN for an invalid Date", () => {
		expect(Number.isNaN(dateToSerial(new Date(Number.NaN)))).toBe(true);
	});
});
