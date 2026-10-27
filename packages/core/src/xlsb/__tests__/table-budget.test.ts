import { describe, expect, it } from "vitest";
import { isDateFormatCode } from "../../ooxml";
import {
	MAX_XLSB_TABLE_ENTRIES,
	parseXlsbStrings,
	parseXlsbStyles,
	parseXlsbWorkbook,
	pushWithinBudget,
} from "../parts";
import { SHEET, SI } from "../records";
import { record, stylesBin, u32, wideString } from "./styles-bin";

// F11.59: the sheet list, shared-string table and cell-format table each fail typed past a budget
// below V8's array limit, where a table of millions of 2-byte records threw a bare RangeError.

const tableError = (what: string) =>
	expect.objectContaining({
		name: "XlsxError",
		code: "part-too-large",
		message: expect.stringContaining(`more than 2 ${what}, the table budget`),
	});
const workbookBin = (count: number) =>
	Uint8Array.from(
		Array.from({ length: count }, (_, i) =>
			record(SHEET, [
				...u32(0),
				...u32(i + 1),
				...wideString(`rId${i + 1}`),
				...wideString(`S${i}`),
			]),
		).flat(),
	);
const stringsBin = (count: number) =>
	Uint8Array.from(
		Array.from({ length: count }, (_, i) => record(SI, [0, ...wideString(`s${i}`)])).flat(),
	);

describe("the workbook table budget", () => {
	it("fails typed past the most sheets", () => {
		expect(parseXlsbWorkbook(workbookBin(2), 2).sheets.map((s) => s.name)).toEqual([
			"S0",
			"S1",
		]);
		expect(() => parseXlsbWorkbook(workbookBin(3), 2)).toThrow(tableError("sheets"));
	});

	it("fails typed past the most shared strings", () => {
		expect(parseXlsbStrings(stringsBin(2), 2)).toEqual(["s0", "s1"]);
		expect(() => parseXlsbStrings(stringsBin(3), 2)).toThrow(tableError("shared strings"));
	});

	it("fails typed past the most cell formats", () => {
		expect(parseXlsbStyles(stylesBin([], [0, 14]), isDateFormatCode, 2).formatCode(1)).toBe(
			"mm-dd-yy",
		);
		expect(() => parseXlsbStyles(stylesBin([], [0, 14, 2]), isDateFormatCode, 2)).toThrow(
			tableError("cell formats"),
		);
	});

	it("sits below V8's 2^27-element array limit", () => {
		expect(MAX_XLSB_TABLE_ENTRIES).toBeLessThan(2 ** 27);
	});
});

describe("pushWithinBudget", () => {
	it("pushes below the budget", () => {
		const list = ["a"];
		pushWithinBudget(list, "b", 2, "shared strings");
		expect(list).toEqual(["a", "b"]);
	});

	it("fails typed at the budget", () => {
		expect(() => pushWithinBudget(["a", "b"], "c", 2, "shared strings")).toThrow(
			tableError("shared strings"),
		);
	});
});
