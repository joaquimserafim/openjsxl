import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F11.10 — the date system is part of the SHARED model: `Workbook.date1904` on read,
// `WorkbookInput.date1904` on write, carried by the bridge. Before this, a 1904 workbook rewrote as a
// 1900 workbook: date-typed cells re-serialized fine, but every raw serial the writer copies verbatim
// (a formula's cached <v>, a validation operand) silently shifted 1,462 days. Now a 1904 workbook
// bridges to a byte-identical 1904 workbook.

const decoder = new TextDecoder();
const workbookXml = async (bytes: Uint8Array) =>
	decoder.decode(await openZip(bytes).read("xl/workbook.xml"));

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		total += value.length;
	}
	const out = new Uint8Array(total);
	let o = 0;
	for (const c of chunks) {
		out.set(c, o);
		o += c.length;
	}
	return out;
}

// A1 is a date; B1 = A1+30 carries its cached result as a RAW General serial (30 under 1904).
const rows = [[new Date(Date.UTC(1904, 0, 1)), { formula: "A1+30", value: 30 }]];

describe("WorkbookInput.date1904 — the date system on the writer side (F11.10)", () => {
	it("selects the 1904 epoch from the input itself, for both writers", async () => {
		const input: WorkbookInput = { sheets: [{ name: "S", rows }], date1904: true };
		const buffered = await writeXlsx(input);
		expect(await workbookXml(buffered)).toContain('<workbookPr date1904="1"/>');
		expect((await openXlsx(buffered)).date1904).toBe(true);
		const streamed = await drain(streamXlsx(input));
		expect(await workbookXml(streamed)).toContain('<workbookPr date1904="1"/>');
		expect((await openXlsx(streamed)).date1904).toBe(true);
	});

	it("an explicit WriteOptions.date1904 overrides the input flag — in both directions", async () => {
		const input: WorkbookInput = { sheets: [{ name: "S", rows }], date1904: true };
		expect((await openXlsx(await writeXlsx(input, { date1904: false }))).date1904).toBe(false);
		const plain: WorkbookInput = { sheets: [{ name: "S", rows }] };
		expect((await openXlsx(await writeXlsx(plain, { date1904: true }))).date1904).toBe(true);
	});

	it("absent or false is the 1900 system and byte-identical to a workbook that never set it", async () => {
		const a = await writeXlsx({ sheets: [{ name: "S", rows }] });
		const b = await writeXlsx({ sheets: [{ name: "S", rows }], date1904: false });
		expect(Buffer.compare(a, b)).toBe(0);
		expect(await workbookXml(a)).not.toContain("date1904");
	});

	it("rejects a non-boolean flag with a typed error", async () => {
		// biome-ignore lint/suspicious/noExplicitAny: exercising a hostile JS caller past the types
		const bad = { sheets: [{ name: "S", rows }], date1904: "yes" as any };
		await expect(writeXlsx(bad)).rejects.toThrow(XlsxError);
		await expect(writeXlsx(bad)).rejects.toMatchObject({ code: "invalid-input" });
	});

	it("the bridge carries it: a 1904 workbook rewrites BYTE-IDENTICAL, cached serials intact", async () => {
		const source = await writeXlsx({ sheets: [{ name: "S", rows }] }, { date1904: true });
		const book = await openXlsx(source);
		expect(book.date1904).toBe(true);
		const input = await workbookToInput(book);
		expect(input.date1904).toBe(true);
		const rewritten = await writeXlsx(input);
		expect(Buffer.compare(rewritten, source)).toBe(0);
		// And the cached serial still means what it meant: 30 days after 1904-01-01, not 1899-12-30+30.
		const again = await openXlsx(rewritten);
		expect(again.date1904).toBe(true);
		expect(again.sheet("S").cell("A1").value).toEqual(new Date(Date.UTC(1904, 0, 1)));
		expect(again.sheet("S").cell("B1").value).toBe(30);
	});

	it("the bridge leaves a 1900 workbook's input untouched (no date1904 key — byte-identity)", async () => {
		const source = await writeXlsx({ sheets: [{ name: "S", rows }] });
		const input = await workbookToInput(await openXlsx(source));
		expect("date1904" in input).toBe(false);
		expect(Buffer.compare(await writeXlsx(input), source)).toBe(0);
	});
});
