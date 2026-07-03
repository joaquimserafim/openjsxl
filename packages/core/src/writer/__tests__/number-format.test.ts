import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { writeXlsx } from "../workbook";

// F4.3 — number-format write. Codes travel as STRINGS (what Excel's Custom dialog shows and what
// numberFormat(ref) returns); ids are file-internal. Exact built-in matches reuse their id with
// no <numFmts> entry; everything else interns from 164 up, deduped, first-encounter order.

const decoder = new TextDecoder();

async function stylesPartOf(bytes: Uint8Array): Promise<string> {
	return decoder.decode(await openZip(bytes).read("xl/styles.xml"));
}

describe("writeXlsx — number formats", () => {
	it("maps an exact built-in code to its id with no <numFmts> entry", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[{ value: 0.42, style: { numberFormat: "0.00%" } }]] }],
		});
		const styles = await stylesPartOf(bytes);
		expect(styles).not.toContain("<numFmts");
		expect(styles).toContain('numFmtId="10"'); // built-in 0.00%
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").numberFormat("A1")).toBe("0.00%");
		expect(wb.sheet("S").style("A1")).toEqual({ numberFormat: "0.00%" });
	});

	it("interns custom codes from 164 up, deduped, in first-encounter order", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ value: 1, style: { numberFormat: '0.0"x"' } },
							{ value: 2, style: { numberFormat: "#,##0.000" } },
							{ value: 3, style: { numberFormat: '0.0"x"' } }, // dupe of the first
						],
					],
				},
			],
		});
		const styles = await stylesPartOf(bytes);
		expect(styles).toContain('<numFmts count="2">');
		expect(styles).toContain('<numFmt numFmtId="164" formatCode="0.0&quot;x&quot;"/>');
		expect(styles).toContain('<numFmt numFmtId="165" formatCode="#,##0.000"/>');

		const sheet = (await openXlsx(bytes)).sheet("S");
		expect(sheet.numberFormat("A1")).toBe('0.0"x"');
		expect(sheet.numberFormat("B1")).toBe("#,##0.000");
		// The dupe shares the first cell's format — and its cached CellStyle object.
		expect(sheet.style("C1")).toBe(sheet.style("A1"));
	});

	it("escapes quotes in format codes and round-trips them verbatim", async () => {
		const code = '"kg" 0.0';
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[{ value: 7, style: { numberFormat: code } }]] }],
			}),
		);
		expect(wb.sheet("S").numberFormat("A1")).toBe(code);
	});

	it("a number with a date-sniffing code re-reads as a date (Excel semantics)", async () => {
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [[{ value: 43831, style: { numberFormat: "yyyy-mm-dd" } }]],
					},
				],
			}),
		);
		const cell = wb.sheet("S").cell("A1");
		expect(cell.type).toBe("date");
		expect((cell.value as Date).toISOString()).toBe("2020-01-01T00:00:00.000Z");
	});

	it("a Date with a user code keeps it — including a NON-date code (re-reads as number)", async () => {
		const date = new Date(Date.UTC(2020, 0, 1));
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							[
								{ value: date, style: { numberFormat: "yyyy/mm/dd hh:mm" } },
								{ value: date, style: { numberFormat: "0.00" } },
							],
						],
					},
				],
			}),
		);
		const sheet = wb.sheet("S");
		// A1: user's date code — still a date, with the user's format visible.
		expect(sheet.cell("A1").type).toBe("date");
		expect(sheet.numberFormat("A1")).toBe("yyyy/mm/dd hh:mm");
		// B1: the caller explicitly chose a numeric format — faithful to Excel, the serial shows
		// as a number on re-read.
		expect(sheet.cell("B1").type).toBe("number");
		expect(sheet.cell("B1").value).toBe(43831);
	});

	it("a bare Date still gets the implicit built-in date format (byte-compat)", async () => {
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[new Date(Date.UTC(2020, 0, 1))]] }] }),
		);
		expect(wb.sheet("S").numberFormat("A1")).toBe("mm-dd-yy");
		expect(wb.sheet("S").cell("A1").type).toBe("date");
	});

	it("'General' is the absence of a format — no styles part at all", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[{ value: 1, style: { numberFormat: "General" } }]] }],
		});
		expect(openZip(bytes).has("xl/styles.xml")).toBe(false);
	});

	it("combines a number format with other components on one xf", async () => {
		const style = { numberFormat: "#,##0.00", font: { bold: true } } as const;
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[{ value: 1234.5, style }]] }] }),
		);
		expect(wb.sheet("S").style("A1")).toEqual(style);
	});

	it('a quoted "[h]" literal in a code does NOT date-flip the value (review regression)', async () => {
		// Pre-fix, the elapsed-time sniff ran on the raw code, so this plain number format
		// re-read as a Date while Excel/openpyxl show a number.
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S",
						rows: [[{ value: 5, style: { numberFormat: '"[h] rate" 0.00' } }]],
					},
				],
			}),
		);
		expect(wb.sheet("S").cell("A1")).toMatchObject({ type: "number", value: 5 });
		expect(wb.sheet("S").numberFormat("A1")).toBe('"[h] rate" 0.00');
	});

	it("tab/newline inside a code survive as character references (review regression)", async () => {
		// Literal whitespace in an ATTRIBUTE is normalized to spaces by conforming parsers; the
		// writer emits &#9;/&#10; so every reader — ours and Excel's — sees the same code.
		const code = '0.0"a\nb\tc"';
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [[{ value: 1, style: { numberFormat: code } }]] }],
		});
		const styles = await stylesPartOf(bytes);
		expect(styles).toContain('formatCode="0.0&quot;a&#10;b&#9;c&quot;"');
		expect((await openXlsx(bytes)).sheet("S").numberFormat("A1")).toBe(code);
	});

	it("rejects a non-string, empty, or XML-unsafe format code", async () => {
		async function code(style: unknown): Promise<string | undefined> {
			const e = await writeXlsx({
				sheets: [{ name: "S", rows: [[{ value: 1, style: style as never }]] }],
			}).then(
				() => undefined,
				(err) => err,
			);
			expect(e).toBeInstanceOf(XlsxError);
			return (e as XlsxError).code;
		}
		expect(await code({ numberFormat: 10 })).toBe("invalid-input");
		expect(await code({ numberFormat: "" })).toBe("invalid-input");
		expect(await code({ numberFormat: `0.0${String.fromCharCode(3)}` })).toBe("invalid-input");
	});
});
