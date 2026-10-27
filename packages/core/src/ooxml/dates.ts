// Excel stores dates and times as serial numbers. Whether a number IS a date is
// decided by the cell's number format (see ./styles), never by the value itself.
// Two epoch systems exist, selected by the workbook's date1904 flag.
//
// 1900 system: serial 1 is 1900-01-01. Excel wrongly treats 1900 as a leap year, so
// serial 60 is a phantom 1900-02-29. Anchoring at 1899-12-30 makes every date from
// 1900-03-01 onward correct — the convention every major library follows.
// 1904 system (legacy Mac): serial 0 is 1904-01-01.

const MS_PER_DAY = 86_400_000;
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30);
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1);

export function serialToDate(serial: number, date1904 = false): Date {
	const epoch = date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC;
	return new Date(epoch + Math.round(serial * MS_PER_DAY));
}

// The exact inverse of serialToDate, for the writer: a JS Date → the Excel serial number that
// serialToDate maps back to the same instant. Because both anchor on the same epoch, any Date the
// writer emits round-trips through the reader losslessly (to the millisecond) — including the
// pre-1900-03-01 range where Excel's phantom leap day makes the *absolute* serial off by one, since
// we invert this implementation, not Excel's calendar. The result carries a fractional day for the
// time-of-day component, exactly as Excel stores it. Returns NaN for an invalid Date (getTime NaN);
// the writer rejects that before emitting.
export function dateToSerial(date: Date, date1904 = false): number {
	const epoch = date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC;
	return (date.getTime() - epoch) / MS_PER_DAY;
}

// ISO 8601 text as a `t="d"` cell stores it (F11.60): YYYY-MM-DD, then optionally THH:MM, :SS, a
// fraction and Z. `\d` matches ASCII digits only, and every field but the fraction has a fixed width.
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?(Z)?$/;
const DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/**
 * Parse an ISO 8601 date or date-time, as a `t="d"` cell stores it, into a UTC wall-clock Date: the
 * strict form `YYYY-MM-DD[THH:MM[:SS[.fraction]]][Z]`, read like a serial date (no time zone). Every
 * field is range-checked: years 0001–9999, real calendar days, hours 0–23, minutes and seconds 0–59.
 * Fraction digits past milliseconds are dropped. Anything else, time-zone offsets and bare times
 * included, gives `undefined`, so no text invents a date (F11.60).
 *
 * @example parseIsoDate("2024-03-15T13:45:30.5Z") // 2024-03-15T13:45:30.500Z
 * @example parseIsoDate("2024-03-15T13:45+01:00") // undefined (not the strict form)
 * @example parseIsoDate("2023-02-29") // undefined (no such day)
 * @example parseIsoDate("2024-03-15T24:00") // undefined (no such hour)
 */
export function parseIsoDate(text: string): Date | undefined {
	const m = ISO_DATE.exec(text);
	if (m === null) return undefined;
	const year = Number(m[1]);
	const month = Number(m[2]);
	const day = Number(m[3]);
	const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
	const monthDays = month === 2 && leap ? 29 : (DAYS_IN_MONTH[month - 1] ?? 0);
	if (year < 1 || day < 1 || day > monthDays) return undefined;
	const hour = Number(m[4] ?? 0);
	const minute = Number(m[5] ?? 0);
	const second = Number(m[6] ?? 0);
	if (hour > 23 || minute > 59 || second > 59) return undefined;
	const ms = Number((m[7] ?? "").slice(0, 3).padEnd(3, "0"));
	const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second, ms));
	// Date.UTC reads years 0–99 as 1900–1999. Set the year back: a year y and 1900 + y agree on
	// leap years for y in 1–99, so the day stays valid.
	if (year < 100) date.setUTCFullYear(year);
	return date;
}
