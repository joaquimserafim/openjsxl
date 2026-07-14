import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { writeXlsx } from "../workbook";

// F3.2 acceptance: a workbook written by writeXlsx must re-read through the reader (openXlsx) with
// the same values, types, and sheet order. These tests are that round-trip proof, plus the input
// validation and the styles-only-when-needed structural guarantee. (An independent-reader check
// with openpyxl is run out-of-band; see the F3.2 notes.)

async function roundTrip(...args: Parameters<typeof writeXlsx>) {
	return openXlsx(await writeXlsx(...args));
}

describe("writeXlsx — value round-trips through openXlsx", () => {
	it("preserves each cell type and value", async () => {
		const date = new Date(Date.UTC(2020, 0, 1));
		const wb = await roundTrip({
			sheets: [
				{
					name: "Data",
					rows: [
						["hello", 42, true],
						[3.14159, -5, false],
						[date, 'a & b < c > d "q"'],
					],
				},
			],
		});
		const sheet = wb.sheet("Data");

		expect(sheet.cell("A1")).toMatchObject({ type: "string", value: "hello" });
		expect(sheet.cell("B1")).toMatchObject({ type: "number", value: 42 });
		expect(sheet.cell("C1")).toMatchObject({ type: "boolean", value: true });
		expect(sheet.cell("A2")).toMatchObject({ type: "number", value: 3.14159 });
		expect(sheet.cell("B2")).toMatchObject({ type: "number", value: -5 });
		expect(sheet.cell("C2")).toMatchObject({ type: "boolean", value: false });

		const a3 = sheet.cell("A3");
		expect(a3.type).toBe("date");
		expect((a3.value as Date).getTime()).toBe(date.getTime());
		// Entities survive the escape → decode round-trip.
		expect(sheet.cell("B3")).toMatchObject({ type: "string", value: 'a & b < c > d "q"' });
	});

	it("omits empty cells (null / undefined / holes) as sparse", async () => {
		const wb = await roundTrip({
			// biome-ignore lint/suspicious/noSparseArray: exercising a genuine array hole
			sheets: [{ name: "S", rows: [["A", null, "C", undefined], [], [, "B2"]] }],
		});
		const sheet = wb.sheet("S");
		expect(sheet.cell("B1").type).toBe("empty");
		expect(sheet.cell("D1").type).toBe("empty");
		expect(sheet.cell("A1").value).toBe("A");
		expect(sheet.cell("C1").value).toBe("C");
		expect(sheet.cell("B3").value).toBe("B2");

		const rows = [];
		for await (const row of sheet.rows()) rows.push(row);
		// Row 2 is entirely empty → absent; row 1 has 2 populated cells; row 3 has 1.
		expect(rows.map((r) => r.index)).toEqual([1, 3]);
		expect(rows[0]?.cells.length).toBe(2);
	});

	it("preserves whitespace-significant strings", async () => {
		const wb = await roundTrip({ sheets: [{ name: "S", rows: [["  padded  "]] }] });
		expect(wb.sheet("S").cell("A1").value).toBe("  padded  ");
	});

	it("keeps legal whitespace (tab/newline) and astral characters (emoji)", async () => {
		const text = "line1\nline2\twith tab";
		const emoji = "café 😀 – 汉字";
		const wb = await roundTrip({ sheets: [{ name: "S", rows: [[text, emoji]] }] });
		expect(wb.sheet("S").cell("A1").value).toBe(text);
		expect(wb.sheet("S").cell("B1").value).toBe(emoji);
	});

	it("preserves multiple sheets in tab order", async () => {
		const wb = await roundTrip({
			sheets: [
				{ name: "First", rows: [[1]] },
				{ name: "Second", rows: [[2]] },
				{ name: "Third", rows: [[3]] },
			],
		});
		expect(wb.sheets.map((s) => s.name)).toEqual(["First", "Second", "Third"]);
		expect(wb.sheet("Second").cell("A1").value).toBe(2);
	});

	it("round-trips dates under the 1904 epoch", async () => {
		const date = new Date(Date.UTC(2020, 0, 1));
		const wb = await roundTrip({ sheets: [{ name: "S", rows: [[date]] }] }, { date1904: true });
		const cell = wb.sheet("S").cell("A1");
		expect(cell.type).toBe("date");
		expect((cell.value as Date).getTime()).toBe(date.getTime());
	});

	it("writes a valid dimension", async () => {
		const wb = await roundTrip({
			sheets: [
				{
					name: "S",
					rows: [
						[1, 2, 3],
						[4, 5, 6],
					],
				},
			],
		});
		expect(wb.sheet("S").dimension).toBe("A1:C2");
	});
});

describe("writeXlsx — structure", () => {
	it("emits styles.xml only when a date is present", async () => {
		const withDate = openZip(
			await writeXlsx({ sheets: [{ name: "S", rows: [[new Date(Date.UTC(2020, 0, 1))]] }] }),
		);
		expect(withDate.has("xl/styles.xml")).toBe(true);

		const noDate = openZip(await writeXlsx({ sheets: [{ name: "S", rows: [["text", 1]] }] }));
		expect(noDate.has("xl/styles.xml")).toBe(false);
	});

	it("is deterministic — identical input yields identical bytes", async () => {
		const input = {
			sheets: [{ name: "S", rows: [["a", 1, new Date(Date.UTC(2021, 5, 15))]] }],
		};
		const first = await writeXlsx(input);
		const second = await writeXlsx(input);
		expect(Array.from(first)).toEqual(Array.from(second));
	});
});

describe("writeXlsx — input validation (invalid-input)", () => {
	async function code(fn: () => Promise<unknown>): Promise<string | undefined> {
		const e = await fn().then(
			() => undefined,
			(err) => err,
		);
		expect(e).toBeInstanceOf(XlsxError);
		return (e as XlsxError).code;
	}

	it("rejects a workbook with no sheets", async () => {
		expect(await code(() => writeXlsx({ sheets: [] }))).toBe("invalid-input");
	});

	it("rejects an empty or over-long sheet name", async () => {
		expect(await code(() => writeXlsx({ sheets: [{ name: "", rows: [] }] }))).toBe(
			"invalid-input",
		);
		expect(await code(() => writeXlsx({ sheets: [{ name: "a".repeat(32), rows: [] }] }))).toBe(
			"invalid-input",
		);
	});

	it("rejects forbidden characters in a sheet name", async () => {
		expect(await code(() => writeXlsx({ sheets: [{ name: "a/b", rows: [] }] }))).toBe(
			"invalid-input",
		);
		expect(await code(() => writeXlsx({ sheets: [{ name: "a[b]", rows: [] }] }))).toBe(
			"invalid-input",
		);
	});

	it("rejects duplicate sheet names case-insensitively", async () => {
		expect(
			await code(() =>
				writeXlsx({
					sheets: [
						{ name: "Data", rows: [] },
						{ name: "data", rows: [] },
					],
				}),
			),
		).toBe("invalid-input");
	});

	it("rejects non-finite numbers and invalid dates", async () => {
		expect(await code(() => writeXlsx({ sheets: [{ name: "S", rows: [[Number.NaN]] }] }))).toBe(
			"invalid-input",
		);
		expect(
			await code(() =>
				writeXlsx({ sheets: [{ name: "S", rows: [[Number.POSITIVE_INFINITY]] }] }),
			),
		).toBe("invalid-input");
		expect(
			await code(() =>
				writeXlsx({ sheets: [{ name: "S", rows: [[new Date(Number.NaN)]] }] }),
			),
		).toBe("invalid-input");
	});

	it("rejects an unsupported cell value type", async () => {
		expect(
			// biome-ignore lint/suspicious/noExplicitAny: passing a bad value a JS caller could
			await code(() => writeXlsx({ sheets: [{ name: "S", rows: [[{} as any]] }] })),
		).toBe("invalid-input");
	});

	it("stores a control character in a cell string via the ST_Xstring escape (F9.6)", async () => {
		// These used to REJECT typed; string content now escapes as _xHHHH_ (the convention Excel
		// decodes), so the value survives write→read — and the raw bytes stay XML-well-formed.
		const nul = String.fromCharCode(0);
		const bel = String.fromCharCode(7);
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[`a${nul}b`, `x${bel}`]] }] });
		const xml = new TextDecoder().decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"));
		expect(xml).toContain("a_x0000_b");
		expect(xml).toContain("x_x0007_");
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").cell("A1").value).toBe(`a${nul}b`);
		expect(wb.sheet("S").cell("B1").value).toBe(`x${bel}`);
	});

	it("rejects a control character in a sheet name", async () => {
		const bel = String.fromCharCode(7);
		expect(
			await code(() => writeXlsx({ sheets: [{ name: `Sheet${bel}`, rows: [[1]] }] })),
		).toBe("invalid-input");
	});

	it("stores a lone surrogate in a cell string via the ST_Xstring escape (F9.6)", async () => {
		// TextEncoder would silently mangle a lone surrogate to U+FFFD — the escape carries the
		// exact code unit instead, so even this survives the trip.
		const loneHigh = String.fromCharCode(0xd800);
		const loneLow = String.fromCharCode(0xdc00);
		const wb = await roundTrip({
			sheets: [{ name: "S", rows: [[`a${loneHigh}`, `${loneLow}b`]] }],
		});
		expect(wb.sheet("S").cell("A1").value).toBe(`a${loneHigh}`);
		expect(wb.sheet("S").cell("B1").value).toBe(`${loneLow}b`);
	});

	it("protects a literal _xHHHH_ so Excel reads back the same text (F9.6)", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["_x0041_"]] }] });
		const xml = new TextDecoder().decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"));
		expect(xml).toContain("_x005F_x0041_"); // the wire form Excel/openpyxl decode to "_x0041_"
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").cell("A1").value).toBe("_x0041_");
	});

	it("rejects a non-array row (string, null) instead of exploding or crashing", async () => {
		// A string row would otherwise be iterated into per-character cells; a null row would throw a
		// raw TypeError. Both must surface as invalid-input.
		// biome-ignore lint/suspicious/noExplicitAny: JS callers can pass a malformed row
		expect(await code(() => writeXlsx({ sheets: [{ name: "S", rows: ["abc" as any] }] }))).toBe(
			"invalid-input",
		);
		// biome-ignore lint/suspicious/noExplicitAny: JS callers can pass a malformed row
		expect(await code(() => writeXlsx({ sheets: [{ name: "S", rows: [null as any] }] }))).toBe(
			"invalid-input",
		);
	});

	it("rejects a non-object workbook (null/undefined/primitive) with a typed error, not a raw throw", async () => {
		// The F9.4 writer fuzzer found that reading `workbook.sheets` off a null/undefined workbook
		// threw a raw TypeError. `requireWorkbookObject` now rejects a non-object workbook as
		// invalid-input, in both the buffered and streaming entries.
		for (const bad of [null, undefined, 42, "x", true]) {
			// biome-ignore lint/suspicious/noExplicitAny: JS callers can pass a non-object workbook
			expect(await code(() => writeXlsx(bad as any))).toBe("invalid-input");
		}
	});
});
