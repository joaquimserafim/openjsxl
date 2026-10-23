import { describe, expect, it } from "vitest";
import {
	CODE_255,
	CODE_256,
	cfSheet,
	styledSheet,
} from "../../ooxml/__tests__/number-format-fixtures";
import { openXlsx } from "../../reader/workbook";
import { streamXlsx, writeXlsx } from "..";

// F11.50: the writer rejects a number-format code longer than Excel's 255 characters with a typed
// error, for cell styles and conditional-formatting dxfs, in both writers. The reader drops such a
// code (see ooxml/__tests__/number-format-limit.test.ts), so the bridge never sees one.

const limitError = expect.objectContaining({
	name: "XlsxError",
	code: "invalid-input",
	message: expect.stringContaining("255"),
});

describe("number-format length limit on write", () => {
	it("accepts and round-trips a 255-character cell code", async () => {
		const book = await openXlsx(await writeXlsx({ sheets: [styledSheet(CODE_255)] }));
		expect(book.sheet("S").numberFormat("A1")).toBe(CODE_255);
	});

	it("rejects a 256-character cell code in both writers", async () => {
		await expect(writeXlsx({ sheets: [styledSheet(CODE_256)] })).rejects.toEqual(limitError);
		const streamed = new Response(
			streamXlsx({ sheets: [styledSheet(CODE_256)] }),
		).arrayBuffer();
		await expect(streamed).rejects.toEqual(limitError);
	});

	it("accepts a 255-character dxf code and rejects a 256-character one", async () => {
		const book = await openXlsx(await writeXlsx({ sheets: [cfSheet(CODE_255)] }));
		expect(book.sheet("S").conditionalFormatting[0]?.rules[0]).toMatchObject({
			dxf: { numberFormat: CODE_255 },
		});
		await expect(writeXlsx({ sheets: [cfSheet(CODE_256)] })).rejects.toEqual(limitError);
	});
});
