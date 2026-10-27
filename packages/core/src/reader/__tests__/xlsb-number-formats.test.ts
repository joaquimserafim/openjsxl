import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { formatRef, parseRef } from "../../ooxml";
import { openXlsb } from "../xlsb";

// F11.59: style 0 is no longer stored per cell. Every corpus xlsb sheet must keep the same number
// formats and styles: one default for refs without a stored style, and these per-ref codes, read over
// the dimension plus two rows and columns. The expectations were read from the reader before F11.59.

const expected: Record<string, Record<string, readonly [string | null, Record<string, string>]>> = {
	"equiv.xlsb": { Data: [null, {}] },
	"sheetjs-short-cells-inline.xlsb": {
		Integers: ["General", {}],
		Decimals: ["General", {}],
		Mixed: [
			"General",
			{
				F1: "yyyy-mm-dd",
				G1: "yyyy-mm-dd hh:mm:ss",
				H1: "0.00",
				E2: "0.00",
				I2: "0.000",
				C3: "0.00",
			},
		],
	},
	"sheetjs-short-cells-sst.xlsb": {
		Integers: ["General", {}],
		Decimals: ["General", {}],
		Mixed: [
			"General",
			{
				F1: "yyyy-mm-dd",
				G1: "yyyy-mm-dd hh:mm:ss",
				H1: "0.00",
				E2: "0.00",
				I2: "0.000",
				C3: "0.00",
			},
		],
	},
	"xlsb-basic.xlsb": {
		Sheet1: ["General", { E1: "mm-dd-yy" }],
		Sheet2: ["General", {}],
		Hidden: ["General", {}],
	},
	"xlsb-date1904.xlsb": { Sheet1: ["General", { A1: "mm-dd-yy", B1: "mm-dd-yy" }] },
	"xlsb-dimensions.xlsb": {
		Empty: ["General", {}],
		A1Value: ["General", {}],
		A1Blank: ["General", { A1: "mm-dd-yy" }],
		Ordinary: ["General", {}],
		InvalidMax: ["General", {}],
		ZeroBounds: ["General", { D7: "mm-dd-yy" }],
		RowDefault: ["General", {}],
		ColumnDefault: ["General", {}],
		MixedDefaults: ["General", {}],
	},
};

describe("xlsb number formats across the corpus", () => {
	for (const [fixture, sheets] of Object.entries(expected)) {
		it(`${fixture} keeps its number formats and styles`, async () => {
			const wb = await openXlsb(await loadFixture(fixture));
			expect(wb.sheets.map((s) => s.name)).toEqual(Object.keys(sheets));
			for (const [name, [fallback, codes]] of Object.entries(sheets)) {
				const sheet = wb.sheet(name);
				expect(sheet.numberFormat("XFD1048576") ?? null).toBe(fallback);
				const seen: Record<string, string> = {};
				const [first = "A1", last = first] = (sheet.dimension ?? "A1").split(":");
				const tl = parseRef(first);
				const br = parseRef(last);
				for (let row = tl.row; row <= br.row + 2; row++) {
					for (let col = tl.col; col <= br.col + 2; col++) {
						const ref = formatRef({ row, col });
						expect(sheet.style(ref)).toBeUndefined();
						const code = sheet.numberFormat(ref) ?? null;
						if (code !== fallback && code !== null) seen[ref] = code;
						else expect(code).toBe(fallback);
					}
				}
				expect(seen).toEqual(codes);
			}
		});
	}
});
