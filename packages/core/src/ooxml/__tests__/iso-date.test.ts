import { afterEach, describe, expect, it } from "vitest";
import { parseIsoDate } from "../dates";

// F11.60: the strict ISO 8601 form a t="d" cell stores, read as a UTC wall-clock date.

const iso = (text: string) => parseIsoDate(text)?.toISOString();

describe("parseIsoDate", () => {
	it("reads a date, and a date-time to the minute, second or fraction", () => {
		expect(iso("2024-03-15")).toBe("2024-03-15T00:00:00.000Z");
		expect(iso("2024-03-15T13:45")).toBe("2024-03-15T13:45:00.000Z");
		expect(iso("2024-03-15T13:45:30")).toBe("2024-03-15T13:45:30.000Z");
		expect(iso("2024-03-15T13:45:30.5")).toBe("2024-03-15T13:45:30.500Z");
		expect(iso("2024-03-15T13:45:30.123")).toBe("2024-03-15T13:45:30.123Z");
	});

	it("drops fraction digits past milliseconds, never rounding into the next second", () => {
		expect(iso("2024-03-15T13:45:30.123456")).toBe("2024-03-15T13:45:30.123Z");
		expect(iso("2024-03-15T13:45:30.1235")).toBe("2024-03-15T13:45:30.123Z");
		expect(iso("2024-03-15T13:45:30.0005")).toBe("2024-03-15T13:45:30.000Z");
		expect(iso("2024-12-31T23:59:59.9999999")).toBe("2024-12-31T23:59:59.999Z");
	});

	it("reads Z as the same wall clock, with or without a time", () => {
		expect(iso("2024-03-15Z")).toBe("2024-03-15T00:00:00.000Z");
		expect(iso("2024-03-15T13:45:30Z")).toBe("2024-03-15T13:45:30.000Z");
	});

	it("reads leap days, and the edges of the four-digit years", () => {
		expect(iso("2024-02-29")).toBe("2024-02-29T00:00:00.000Z");
		expect(iso("2000-02-29")).toBe("2000-02-29T00:00:00.000Z");
		expect(iso("0001-01-01T06:00:00")).toBe("0001-01-01T06:00:00.000Z");
		expect(iso("0004-02-29")).toBe("0004-02-29T00:00:00.000Z");
		expect(iso("0099-12-31")).toBe("0099-12-31T00:00:00.000Z");
		expect(iso("9999-12-31T23:59:59.999")).toBe("9999-12-31T23:59:59.999Z");
	});

	it("rejects days, months and times that do not exist", () => {
		for (const text of [
			"2023-02-29",
			"1900-02-29",
			"2024-02-30",
			"2024-04-31",
			"2024-00-10",
			"2024-13-01",
			"2024-03-00",
			"0000-01-01",
			"2024-03-15T24:00",
			"2024-03-15T23:60",
			"2024-03-15T23:59:60",
		]) {
			expect(parseIsoDate(text), text).toBeUndefined();
		}
	});

	it("rejects time-zone offsets, bare times and every other spelling", () => {
		for (const text of [
			"2024-03-15T13:45:30+01:00",
			"2024-03-15T13:45:30-05:00",
			"2024-03-15+01:00",
			"13:45:30",
			"2024-03-15t13:45:30",
			"2024-03-15T13:45:30z",
			"2024-03-15 13:45:30",
			"2024-03-15T",
			"2024-03-15T13",
			"2024-03-15T13:45:",
			"2024-03-15T13:45:30.",
			"2024-3-15",
			"24-03-15",
			"+2024-03-15",
			"12024-03-15",
			" 2024-03-15",
			"2024-03-15 ",
			"2024-03-15x",
			"２０２４-03-15",
			"45366",
			"",
		]) {
			expect(parseIsoDate(text), JSON.stringify(text)).toBeUndefined();
		}
	});
});

describe("parseIsoDate outside UTC", () => {
	// The reader's dates are UTC wall clocks: the process's time zone must not move them. Node reads
	// TZ again when it changes, so each zone is set for one test and then restored.
	const original = process.env.TZ;
	afterEach(() => {
		if (original === undefined) Reflect.deleteProperty(process.env, "TZ");
		else process.env.TZ = original;
	});

	for (const zone of [
		"America/New_York",
		"Asia/Kolkata",
		"Pacific/Kiritimati",
		"Europe/London",
	]) {
		it(`builds the same instant under TZ=${zone}`, () => {
			process.env.TZ = zone;
			expect(iso("2024-03-15T13:45:30")).toBe("2024-03-15T13:45:30.000Z");
			// 02:30 on 2024-03-10 does not exist in New York (the clocks jump an hour).
			expect(iso("2024-03-10T02:30")).toBe("2024-03-10T02:30:00.000Z");
			expect(iso("0001-01-01T06:00:00")).toBe("0001-01-01T06:00:00.000Z");
		});
	}
});
