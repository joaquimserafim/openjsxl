import { describe, expect, it } from "vitest";
import { decodeCell } from "../cell";
import { dateToSerial } from "../dates";
import type { StyleTable } from "../styles";

// F11.60: a t="d" cell holds ISO 8601 text, read as a date whatever its style or the date system.

const ctx = { sharedStrings: [] };
const d = (value: string | undefined, style?: number) =>
	decodeCell({ ref: "A1", type: "d", value, style }, ctx);

describe('decodeCell — t="d"', () => {
	it("reads ISO 8601 text as a date", () => {
		expect(d("2024-03-15T13:45:30")).toEqual({
			ref: "A1",
			type: "date",
			value: new Date(Date.UTC(2024, 2, 15, 13, 45, 30)),
		});
	});

	it("ignores the cell's style and the 1904 date system", () => {
		const styles: StyleTable = {
			isDateStyle: () => false,
			formatCode: () => "0.00",
			cellStyle: () => undefined,
		};
		const expected = new Date(Date.UTC(2024, 2, 15));
		for (const context of [
			{ ...ctx, styles },
			{ ...ctx, date1904: true },
		]) {
			const cell = decodeCell(
				{ ref: "B2", type: "d", value: "2024-03-15", style: 3 },
				context,
			);
			expect(cell).toEqual({ ref: "B2", type: "date", value: expected });
		}
	});

	it("reads a date before 1900-03-01 exactly, under the reader's 1899-12-30 serial anchor", () => {
		const date = new Date(Date.UTC(1900, 0, 15));
		expect(d("1900-01-15")).toEqual({ ref: "A1", type: "date", value: date });
		// Excel numbers 1900-01-15 as 15, counting a 1900-02-29 that never was. The reader's serials
		// are one higher before 1900-03-01, so the evaluator and the bridge see 16: the known offset
		// that F11.78 documents. t="d" text is one more way such a date arrives.
		expect(dateToSerial(date)).toBe(16);
	});

	it("reads text that is not a real date as empty, never a guessed date", () => {
		for (const value of [
			"2024-02-30",
			"2024-03-15T13:45+01:00",
			"13:45:30",
			"45366",
			"",
			undefined,
		]) {
			expect(d(value), String(value)).toEqual({ ref: "A1", type: "empty", value: null });
		}
	});
});
