import { describe, expect, it } from "vitest";
import { dateFn, datePart, days, monthShift, weekday } from "../builtins";
import type { EvalContext } from "../functions";
import { errorValue } from "../value";

// F11.33: DATE, WEEKDAY, DAYS, the date-part functions, EDATE and EOMONTH, each on every exit, in a
// context without the flag (the 1900 system) and one that sets it (the 1904 system).

const plain: EvalContext = { now: () => 0, random: () => 0 };
const e1904: EvalContext = { ...plain, date1904: true };
const year = (dt: Date) => dt.getUTCFullYear();

describe("dateFn", () => {
	it("returns an argument's error", () => {
		expect(dateFn([errorValue("#N/A"), 1, 1], plain)).toEqual(errorValue("#N/A"));
		expect(dateFn([2020, "one", 1], plain)).toEqual(errorValue("#VALUE!"));
		expect(dateFn([2020, 1, "one"], plain)).toEqual(errorValue("#VALUE!"));
	});

	it("is #NUM! for a year outside 0 to 9999", () => {
		expect(dateFn([10000, 1, 1], plain)).toEqual(errorValue("#NUM!"));
		expect(dateFn([-1, 1, 1], plain)).toEqual(errorValue("#NUM!"));
	});

	it("is #NUM! for a date outside the system's serials", () => {
		expect(dateFn([1903, 12, 31], e1904)).toEqual(errorValue("#NUM!"));
		expect(dateFn([1903, 12, 31], plain)).toBe(1461);
		expect(dateFn([0, 1, -2], plain)).toEqual(errorValue("#NUM!"));
		// 9999-12-32 rolls into 10000-01-01, one day past the last valid serial in each system.
		expect(dateFn([9999, 12, 32], plain)).toEqual(errorValue("#NUM!"));
		expect(dateFn([9999, 12, 32], e1904)).toEqual(errorValue("#NUM!"));
	});

	it("gives the serial in the context's system", () => {
		expect(dateFn([1904, 1, 1], plain)).toBe(1462);
		expect(dateFn([1904, 1, 1], e1904)).toBe(0);
		expect(dateFn([9999, 12, 31], plain)).toBe(2958465);
		expect(dateFn([9999, 12, 31], e1904)).toBe(2957003);
	});
});

describe("weekday", () => {
	it("returns the serial's or the type's error", () => {
		expect(weekday([-1], plain)).toEqual(errorValue("#NUM!"));
		expect(weekday([0, "one"], plain)).toEqual(errorValue("#VALUE!"));
		expect(weekday([2957004], e1904)).toEqual(errorValue("#NUM!"));
	});

	it("numbers the day by return_type in the context's system", () => {
		expect(weekday([0], plain)).toBe(7); // 1899-12-30, a Saturday
		expect(weekday([0], e1904)).toBe(6); // 1904-01-01, a Friday
		expect(weekday([0, 2], e1904)).toBe(5);
		expect(weekday([0, 3], e1904)).toBe(4);
		expect(weekday([0, 11], e1904)).toBe(5);
		expect(weekday([0, 17], e1904)).toBe(6);
	});

	it("is #NUM! for any other return_type", () => {
		expect(weekday([0, 4], plain)).toEqual(errorValue("#NUM!"));
		expect(weekday([0, 18], e1904)).toEqual(errorValue("#NUM!"));
	});
});

describe("days", () => {
	it("returns the end date's or the start date's error", () => {
		expect(days([errorValue("#N/A"), 0], plain)).toEqual(errorValue("#N/A"));
		expect(days([10, -1], plain)).toEqual(errorValue("#NUM!"));
		expect(days([2957004, 0], e1904)).toEqual(errorValue("#NUM!"));
		expect(days([0, 2957004], e1904)).toEqual(errorValue("#NUM!"));
	});

	it("gives the whole days between them", () => {
		expect(days([2957004, 4], plain)).toBe(2957000);
		expect(days([10.9, 2.5], e1904)).toBe(8);
	});
});

describe("datePart", () => {
	it("returns the serial's error", () => {
		expect(datePart([-1], plain, year)).toEqual(errorValue("#NUM!"));
		expect(datePart([2957004], e1904, year)).toEqual(errorValue("#NUM!"));
	});

	it("reads the part in the context's system", () => {
		expect(datePart([0], plain, year)).toBe(1899);
		expect(datePart([0], e1904, year)).toBe(1904);
		expect(datePart([2958465], plain, year)).toBe(9999);
	});
});

describe("monthShift", () => {
	it("returns the serial's or the months' error", () => {
		expect(monthShift([-1, 1], plain, "same")).toEqual(errorValue("#NUM!"));
		expect(monthShift([0, "one"], plain, "same")).toEqual(errorValue("#VALUE!"));
		expect(monthShift([2957004, -1], e1904, "same")).toEqual(errorValue("#NUM!"));
	});

	it("is #NUM! for a result outside the system's serials", () => {
		expect(monthShift([0, -1], e1904, "same")).toEqual(errorValue("#NUM!"));
		expect(monthShift([0, -1], e1904, "end")).toEqual(errorValue("#NUM!"));
		expect(monthShift([2957003, 1], e1904, "same")).toEqual(errorValue("#NUM!"));
		expect(monthShift([2957003, 1], plain, "same")).toBe(2957034);
	});

	it("moves to the same day or the month's end in the context's system", () => {
		expect(monthShift([0, 1], plain, "same")).toBe(31); // 1900-01-30
		expect(monthShift([0, 1], e1904, "same")).toBe(31); // 1904-02-01
		expect(monthShift([0, 0], plain, "end")).toBe(1); // 1899-12-31
		expect(monthShift([0, 0], e1904, "end")).toBe(30); // 1904-01-31
		expect(monthShift([30, 1], e1904, "same")).toBe(59); // 1904-01-31 to 1904-02-29
	});
});
