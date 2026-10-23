import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { parseStyles } from "../styles";
import {
	CODE_255,
	CODE_256,
	cfSheet,
	corpusFormatCodes,
	DATE_255,
	DATE_256,
	legacyIsDate,
	legacyStrip,
	randomCodes,
	styledSheet,
	stylesXml,
	workbookWithFormat,
} from "./number-format-fixtures";

describe("number-format fixtures", () => {
	it("sizes the codes at and one past the limit", () => {
		expect([DATE_255.length, DATE_256.length, CODE_255.length, CODE_256.length]).toEqual([
			255, 256, 255, 256,
		]);
	});

	it("keeps the 1.0.0 oracles' known answers", () => {
		expect(legacyStrip("[Red]0.00")).toBe("0.00");
		expect(legacyStrip("0[x")).toBe("0[x");
		expect([legacyIsDate("yyyy-mm-dd"), legacyIsDate("[h]"), legacyIsDate("0.00")]).toEqual([
			true,
			true,
			false,
		]);
	});

	it("generates the same random codes on every run", () => {
		expect([...randomCodes(50)]).toEqual([...randomCodes(50)]);
		expect([...randomCodes(5)]).toHaveLength(5);
	});

	it("samples a non-empty corpus", async () => {
		expect((await corpusFormatCodes()).length).toBeGreaterThan(0);
	});

	it("builds a styles part the reader parses", () => {
		const styles = parseStyles(stylesXml([[164, "0.0"]], [0, 164]));
		expect(styles.formatCode(1)).toBe("0.0");
	});

	it("replaces the seeded code in a written workbook", async () => {
		const book = await openXlsx(await workbookWithFormat("0.000"));
		expect(book.sheet("S").numberFormat("A1")).toBe("0.000");
	});

	it("builds styled and conditional-formatting sheets", () => {
		expect(styledSheet("0.00").rows).toEqual([
			[{ value: 1.5, style: { numberFormat: "0.00" } }],
		]);
		expect(cfSheet("0.0%").conditionalFormatting?.[0]?.rules[0]).toMatchObject({
			dxf: { numberFormat: "0.0%" },
		});
	});
});
