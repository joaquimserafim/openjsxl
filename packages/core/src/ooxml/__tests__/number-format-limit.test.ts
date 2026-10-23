import { describe, expect, it } from "vitest";
import { openXlsx, streamSheetRows } from "../../reader/workbook";
import { parseDxfs } from "../dxf";
import { parseStyles } from "../styles";
import { DATE_255, DATE_256, stylesXml, workbookWithFormat } from "./number-format-fixtures";

// F11.50: Excel caps a number-format code at 255 characters. The reader drops a longer code to
// absent (no format, never a date, no built-in fallback); the writer rejects one (see the writer's
// number-format-limit test). The shared bound is MAX_NUMBER_FORMAT_LEN.

describe("number-format length limit on read (styles.xml)", () => {
	it("keeps a 255-character code", () => {
		const styles = parseStyles(stylesXml([[164, DATE_255]], [0, 164]));
		expect(styles.formatCode(1)).toBe(DATE_255);
		expect(styles.isDateStyle(1)).toBe(true);
		expect(styles.cellStyle(1)?.numberFormat).toBe(DATE_255);
	});

	it("drops a 256-character code to absent", () => {
		const styles = parseStyles(stylesXml([[164, DATE_256]], [0, 164]));
		expect(styles.formatCode(1)).toBeUndefined();
		expect(styles.isDateStyle(1)).toBe(false);
		expect(styles.cellStyle(1)?.numberFormat).toBeUndefined();
	});

	it("drops an oversized redefinition of a built-in id without falling back to the built-in", () => {
		const styles = parseStyles(stylesXml([[14, DATE_256]], [0, 14]));
		expect(styles.formatCode(1)).toBeUndefined();
		expect(styles.isDateStyle(1)).toBe(false);
	});
});

describe("number-format length limit on read (dxfs)", () => {
	it("keeps a 255-character dxf code and drops a 256-character one", () => {
		const dxfs = parseDxfs(
			'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
				'<dxfs count="2">' +
				`<dxf><numFmt numFmtId="170" formatCode="${DATE_255}"/></dxf>` +
				`<dxf><numFmt numFmtId="171" formatCode="${DATE_256}"/></dxf>` +
				"</dxfs></styleSheet>",
		);
		expect(dxfs.map((dxf) => dxf.numberFormat)).toEqual([DATE_255, undefined]);
	});
});

describe("number-format length limit through the readers", () => {
	it("types a cell with a 255-character date code as a date", async () => {
		const book = await openXlsx(await workbookWithFormat(DATE_255));
		expect(book.sheet("S").numberFormat("A1")).toBe(DATE_255);
		expect(book.sheet("S").cell("A1").type).toBe("date");
	});

	it("reads a cell with an oversized code as a plain number in both readers", async () => {
		for (const code of [DATE_256, "[".repeat(20_000)]) {
			const bytes = await workbookWithFormat(code);
			const book = await openXlsx(bytes);
			expect(book.sheet("S").numberFormat("A1")).toBeUndefined();
			expect(book.sheet("S").cell("A1")).toMatchObject({ type: "number", value: 45000 });
			const streamed: unknown[] = [];
			for await (const row of streamSheetRows(bytes)) streamed.push(...row.cells);
			expect(streamed).toEqual([expect.objectContaining({ type: "number", value: 45000 })]);
		}
	});
});
