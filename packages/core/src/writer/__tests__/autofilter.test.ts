import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F10.2 — sheet-level autoFilter: write + the paired hidden _xlnm._FilterDatabase name + bridge carry.
// A filtered sheet emits <autoFilter> in the CT_Worksheet slot after </sheetData>, synthesizes the
// _FilterDatabase name (which the reader strips), and re-reads through Worksheet.autoFilter; an
// unfiltered sheet keeps its exact pre-F10.2 bytes.

const decoder = new TextDecoder();
const sheetXmlOf = async (bytes: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"));
const workbookXmlOf = async (bytes: Uint8Array): Promise<string> =>
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

async function writeErr(wb: unknown): Promise<XlsxError> {
	try {
		await writeXlsx(wb as WorkbookInput);
	} catch (e) {
		if (e instanceof XlsxError) return e;
		throw e;
	}
	throw new Error("expected writeXlsx to reject");
}

describe("writeXlsx — <autoFilter> emission", () => {
	it("emits <autoFilter> in the slot right after </sheetData>, before <mergeCells>", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "Data",
					rows: [["a", "b"]],
					autoFilter: { ref: "A1:B1" },
					merges: ["A2:B2"],
				},
			],
		});
		const xml = await sheetXmlOf(bytes);
		expect(xml).toContain('</sheetData><autoFilter ref="A1:B1"/><mergeCells');
	});

	it("synthesizes a hidden, sheet-scoped _xlnm._FilterDatabase (sheet-qualified absolute range)", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Data", rows: [["a"]], autoFilter: { ref: "A1:C3" } }],
		});
		expect(await workbookXmlOf(bytes)).toContain(
			`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Data'!$A$1:$C$3</definedName>`,
		);
	});

	it("scopes each sheet's _FilterDatabase to its own index; quotes an apostrophe in the sheet name", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{ name: "Plain", rows: [["a"]] },
				{ name: "Bob's", rows: [["a", "b"]], autoFilter: { ref: "A1:B2" } },
			],
		});
		expect(await workbookXmlOf(bytes)).toContain(
			`<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'Bob''s'!$A$1:$B$2</definedName>`,
		);
	});
});

describe("writeXlsx — byte-identity when no autoFilter is used", () => {
	it("an unfiltered sheet emits no <autoFilter> and no <definedNames>", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "Data", rows: [["a", 1]] }] });
		expect(await sheetXmlOf(bytes)).not.toContain("autoFilter");
		expect(await workbookXmlOf(bytes)).not.toContain("definedNames");
	});
});

describe("writeXlsx — round-trip through the reader (shared model)", () => {
	it("re-reads the filter range; _FilterDatabase never surfaces as a defined name", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Data", rows: [["a"]], autoFilter: { ref: "A1:B4" } }],
		});
		const wb = await openXlsx(bytes);
		expect(wb.sheet("Data").autoFilter).toEqual({ ref: "A1:B4" });
		expect(wb.definedNames).toEqual([]);
	});

	it("carries autoFilter through the bridge (read → workbookToInput → write)", async () => {
		const first = await writeXlsx({
			sheets: [{ name: "Data", rows: [["a"]], autoFilter: { ref: "A1:C3" } }],
		});
		const round = await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(first))));
		expect(round.sheet("Data").autoFilter).toEqual({ ref: "A1:C3" });
		expect(round.definedNames).toEqual([]);
	});

	it("keeps a real defined name distinct from the filter (both survive)", async () => {
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [{ name: "Data", rows: [["a"]], autoFilter: { ref: "A1:A1" } }],
				definedNames: [{ name: "MyName", refersTo: "Data!$A$1" }],
			}),
		);
		expect(wb.sheet("Data").autoFilter).toEqual({ ref: "A1:A1" });
		expect(wb.definedNames).toEqual([{ name: "MyName", refersTo: "Data!$A$1" }]);
	});

	it("streamXlsx emits the same filter (reader-equivalent to writeXlsx)", async () => {
		const wb = await openXlsx(
			await drain(
				streamXlsx({
					sheets: [{ name: "Data", rows: [["a"]], autoFilter: { ref: "A1:B2" } }],
				}),
			),
		);
		expect(wb.sheet("Data").autoFilter).toEqual({ ref: "A1:B2" });
		expect(wb.definedNames).toEqual([]);
	});
});

describe("writeXlsx — validation rejects (typed invalid-input)", () => {
	const filtered = (autoFilter: unknown): unknown => ({
		sheets: [{ name: "Data", rows: [["a"]], autoFilter }],
	});

	it("rejects a non-object autoFilter and an unknown property", async () => {
		expect((await writeErr(filtered("A1:B2"))).message).toContain(
			"autoFilter must be an object",
		);
		expect((await writeErr(filtered({ ref: "A1:B2", sort: true }))).message).toContain(
			'unknown property "sort"',
		);
	});

	it("rejects a non-string ref and a non-canonical / out-of-grid ref", async () => {
		expect((await writeErr(filtered({ ref: 5 }))).message).toContain(
			"autoFilter.ref must be a string",
		);
		expect((await writeErr(filtered({ ref: "a1:b2" }))).message).toContain(
			"canonical A1 range",
		);
		expect((await writeErr(filtered({ ref: "A1:ZZZZ9" }))).message).toContain(
			"canonical A1 range",
		);
		expect((await writeErr(filtered({ ref: "A1048577" }))).message).toContain(
			"canonical A1 range",
		);
	});

	it("rejects a caller-supplied _xlnm._FilterDatabase in definedNames (managed by autoFilter)", async () => {
		const err = await writeErr({
			sheets: [{ name: "Data", rows: [["a"]] }],
			definedNames: [{ name: "_xlnm._FilterDatabase", refersTo: "Data!$A$1" }],
		});
		expect(err.code).toBe("invalid-input");
		expect(err.message).toContain("SheetInput.autoFilter");
	});
});
