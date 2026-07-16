import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F10.4 — print setup: printOptions, pageMargins, pageSetup, headerFooter. Emitted in schema order
// between <hyperlinks> and <drawing>; an unset sheet keeps its exact pre-F10.4 bytes.

const decoder = new TextDecoder();
const sheetXmlOf = async (b: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(b).read("xl/worksheets/sheet1.xml"));

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

async function writeErr(wb: unknown): Promise<XlsxError> {
	try {
		await writeXlsx(wb as WorkbookInput);
	} catch (e) {
		if (e instanceof XlsxError) return e;
		throw e;
	}
	throw new Error("expected writeXlsx to reject");
}

describe("writeXlsx — print-setup emission", () => {
	it("emits printOptions → pageMargins → pageSetup → headerFooter, in order after hyperlinks", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "R",
					rows: [["a"]],
					printOptions: { horizontalCentered: true, gridLines: true },
					pageMargins: {
						left: 0.7,
						right: 0.7,
						top: 1,
						bottom: 1,
						header: 0.3,
						footer: 0.3,
					},
					pageSetup: {
						orientation: "landscape",
						paperSize: 9,
						fitToWidth: 1,
						fitToHeight: 0,
					},
					headerFooter: { oddHeader: "&CTitle" },
				},
			],
		});
		const xml = await sheetXmlOf(bytes);
		expect(xml).toContain(
			'<printOptions horizontalCentered="1" gridLines="1"/>' +
				'<pageMargins left="0.7" right="0.7" top="1" bottom="1" header="0.3" footer="0.3"/>' +
				'<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="landscape"/>' +
				"<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>",
		);
	});

	it("escapes & codes in header/footer text", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "R", rows: [["a"]], headerFooter: { oddFooter: "&RPage &P of &N" } }],
		});
		expect(await sheetXmlOf(bytes)).toContain(
			"<oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter>",
		);
	});
});

describe("writeXlsx — byte-identity when no print setup is used", () => {
	it("a sheet without print setup emits none of the four elements", async () => {
		const xml = await sheetXmlOf(
			await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1]] }] }),
		);
		for (const tag of ["printOptions", "pageMargins", "pageSetup", "headerFooter"]) {
			expect(xml).not.toContain(tag);
		}
	});
});

describe("writeXlsx — round-trip through the reader (shared model)", () => {
	it("re-reads all four elements", async () => {
		const input = {
			name: "R",
			rows: [["a"]],
			printOptions: { verticalCentered: true },
			pageMargins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
			pageSetup: { orientation: "portrait", scale: 90, draft: true },
			headerFooter: { oddHeader: "&LLeft", differentOddEven: true, evenHeader: "&REven" },
		} as const;
		const wb = await openXlsx(await writeXlsx({ sheets: [input] }));
		const s = wb.sheet("R");
		expect(s.printOptions).toEqual({ verticalCentered: true });
		expect(s.pageMargins).toEqual(input.pageMargins);
		expect(s.pageSetup).toEqual({ orientation: "portrait", scale: 90, draft: true });
		expect(s.headerFooter).toEqual({
			oddHeader: "&LLeft",
			differentOddEven: true,
			evenHeader: "&REven",
		});
	});

	it("carries print setup through the bridge", async () => {
		const first = await writeXlsx({
			sheets: [
				{
					name: "R",
					rows: [["a"]],
					pageSetup: { orientation: "landscape", fitToWidth: 2 },
					headerFooter: { oddFooter: "&CFooter" },
				},
			],
		});
		const s = (
			await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(first))))
		).sheet("R");
		expect(s.pageSetup).toEqual({ orientation: "landscape", fitToWidth: 2 });
		expect(s.headerFooter).toEqual({ oddFooter: "&CFooter" });
	});

	it("streamXlsx emits the same print setup", async () => {
		const wb = await openXlsx(
			await drain(
				streamXlsx({
					sheets: [{ name: "R", rows: [["a"]], pageSetup: { orientation: "landscape" } }],
				}),
			),
		);
		expect(wb.sheet("R").pageSetup).toEqual({ orientation: "landscape" });
	});
});

describe("writeXlsx — validation rejects (typed invalid-input)", () => {
	const on = (field: string, value: unknown): unknown => ({
		sheets: [{ name: "R", rows: [["a"]], [field]: value }],
	});

	it("rejects unknown properties and bad types", async () => {
		expect((await writeErr(on("printOptions", { nope: true }))).message).toContain(
			'printOptions has unknown property "nope"',
		);
		expect((await writeErr(on("printOptions", { gridLines: 1 }))).message).toContain(
			"printOptions.gridLines must be a boolean",
		);
		expect((await writeErr(on("pageSetup", { orientation: "sideways" }))).message).toContain(
			"pageSetup.orientation must be one of",
		);
		expect((await writeErr(on("headerFooter", { oddHeader: 5 }))).message).toContain(
			"headerFooter.oddHeader must be a string",
		);
	});

	it("rejects a non-finite / out-of-range margin and a partial pageMargins", async () => {
		expect(
			(
				await writeErr(
					on("pageMargins", {
						left: Number.POSITIVE_INFINITY,
						right: 1,
						top: 1,
						bottom: 1,
						header: 1,
						footer: 1,
					}),
				)
			).message,
		).toContain("pageMargins.left must be a finite number");
		// A missing required margin surfaces typed (undefined fails the finite-number check).
		expect((await writeErr(on("pageMargins", { left: 1 }))).message).toContain(
			"pageMargins.right must be a finite number",
		);
	});

	it("rejects a scale out of 10..400 and a uint attr out of range", async () => {
		expect((await writeErr(on("pageSetup", { scale: 9 }))).message).toContain(
			"pageSetup.scale must be an integer in 10..400",
		);
		expect((await writeErr(on("pageSetup", { paperSize: 1e21 }))).message).toContain(
			"pageSetup.paperSize must be an integer in 0..",
		);
	});
});
