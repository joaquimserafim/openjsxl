// Excel stores dates and times as serial numbers. Whether a number IS a date is
// decided by the cell's number format (see ./styles), never by the value itself.
// Two epoch systems exist, selected by the workbook's date1904 flag.
//
// 1900 system: serial 1 is 1900-01-01. Excel wrongly treats 1900 as a leap year, so
// serial 60 is a phantom 1900-02-29. Anchoring at 1899-12-30 makes every date from
// 1900-03-01 onward correct — the convention every major library follows.
// 1904 system (legacy Mac): serial 0 is 1904-01-01.

const MS_PER_DAY = 86_400_000
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30)
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1)

export function serialToDate(serial: number, date1904 = false): Date {
	const epoch = date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC
	return new Date(epoch + Math.round(serial * MS_PER_DAY))
}

// The exact inverse of serialToDate, for the writer: a JS Date → the Excel serial number that
// serialToDate maps back to the same instant. Because both anchor on the same epoch, any Date the
// writer emits round-trips through the reader losslessly (to the millisecond) — including the
// pre-1900-03-01 range where Excel's phantom leap day makes the *absolute* serial off by one, since
// we invert this implementation, not Excel's calendar. The result carries a fractional day for the
// time-of-day component, exactly as Excel stores it. Returns NaN for an invalid Date (getTime NaN);
// the writer rejects that before emitting.
export function dateToSerial(date: Date, date1904 = false): number {
	const epoch = date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC
	return (date.getTime() - epoch) / MS_PER_DAY
}
