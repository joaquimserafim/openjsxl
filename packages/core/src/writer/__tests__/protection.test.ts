import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F10.3 — protection: workbook <workbookProtection>, sheet <sheetProtection>, and per-cell locked/hidden
// (CellStyle.protection). Everything written re-reads through the model; an unprotected workbook keeps
// its exact pre-F10.3 bytes; password material is carried verbatim, never computed.

const decoder = new TextDecoder();
const sheetXmlOf = async (b: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(b).read("xl/worksheets/sheet1.xml"));
const workbookXmlOf = async (b: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(b).read("xl/workbook.xml"));
const stylesXmlOf = async (b: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(b).read("xl/styles.xml"));

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

describe("writeXlsx — <sheetProtection> emission", () => {
	it("emits in the slot after </sheetData>, before <autoFilter>; booleans as 1/0; password verbatim", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "Data",
					rows: [["a"]],
					protection: { sheet: true, formatCells: false, sort: true, password: "C258" },
					autoFilter: { ref: "A1:A1" },
				},
			],
		});
		const xml = await sheetXmlOf(bytes);
		expect(xml).toContain(
			'</sheetData><sheetProtection sheet="1" formatCells="0" sort="1" password="C258"/><autoFilter',
		);
	});

	it("carries the modern hash attributes verbatim", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["a"]],
					protection: {
						sheet: true,
						algorithmName: "SHA-512",
						hashValue: "abc==",
						saltValue: "xyz==",
						spinCount: 100000,
					},
				},
			],
		});
		expect(await sheetXmlOf(bytes)).toContain(
			'<sheetProtection sheet="1" algorithmName="SHA-512" hashValue="abc==" saltValue="xyz==" spinCount="100000"/>',
		);
	});
});

describe("writeXlsx — <workbookProtection> emission", () => {
	it("emits after <workbookPr>, before <sheets>", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [["a"]] }],
			protection: { lockStructure: true, lockWindows: false },
		});
		expect(await workbookXmlOf(bytes)).toContain(
			'<workbookProtection lockStructure="1" lockWindows="0"/><sheets>',
		);
	});
});

describe("writeXlsx — cell-level <protection>", () => {
	it("emits applyProtection + <protection> for an unlocked cell; alignment then protection order", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [
						[
							{
								value: "x",
								style: {
									alignment: { horizontal: "center" },
									protection: { locked: false },
								},
							},
						],
					],
				},
			],
		});
		const xml = await stylesXmlOf(bytes);
		expect(xml).toContain('applyAlignment="1" applyProtection="1">');
		expect(xml).toContain('<alignment horizontal="center"/><protection locked="0"/></xf>');
	});

	it("interns a protection-only style (and a styled BLANK carries it)", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{ name: "S", rows: [[{ value: null, style: { protection: { hidden: true } } }]] },
			],
		});
		expect(await stylesXmlOf(bytes)).toContain(
			'applyProtection="1"><protection hidden="1"/></xf>',
		);
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").style("A1")?.protection).toEqual({ hidden: true });
	});
});

describe("writeXlsx — byte-identity + round-trip", () => {
	it("no protection anywhere → no <sheetProtection>/<workbookProtection>/<protection>", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1]] }] });
		expect(await sheetXmlOf(bytes)).not.toContain("sheetProtection");
		expect(await workbookXmlOf(bytes)).not.toContain("workbookProtection");
	});

	it("round-trips sheet + workbook + cell protection through the bridge", async () => {
		const first = await writeXlsx({
			sheets: [
				{
					name: "Data",
					rows: [["h", { value: "e", style: { protection: { locked: false } } }]],
					protection: { sheet: true, formatCells: false, password: "C258" },
				},
			],
			protection: { lockStructure: true },
		});
		const wb = await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(first))));
		expect(wb.sheet("Data").protection).toEqual({
			sheet: true,
			formatCells: false,
			password: "C258",
		});
		expect(wb.protection).toEqual({ lockStructure: true });
		expect(wb.sheet("Data").style("B1")?.protection).toEqual({ locked: false });
	});

	it("streamXlsx emits the same sheet + workbook protection", async () => {
		const wb = await openXlsx(
			await drain(
				streamXlsx({
					sheets: [{ name: "S", rows: [["a"]], protection: { sheet: true } }],
					protection: { lockStructure: true },
				}),
			),
		);
		expect(wb.sheet("S").protection).toEqual({ sheet: true });
		expect(wb.protection).toEqual({ lockStructure: true });
	});
});

describe("writeXlsx — validation rejects (typed invalid-input)", () => {
	it("rejects a non-boolean sheet flag, an unknown attr, and a bad spinCount", async () => {
		expect(
			(await writeErr({ sheets: [{ name: "S", rows: [["a"]], protection: { sheet: "y" } }] }))
				.message,
		).toContain("protection.sheet must be a boolean");
		expect(
			(await writeErr({ sheets: [{ name: "S", rows: [["a"]], protection: { nope: true } }] }))
				.message,
		).toContain('unknown property "nope"');
		expect(
			(
				await writeErr({
					sheets: [
						{ name: "S", rows: [["a"]], protection: { sheet: true, spinCount: -1 } },
					],
				})
			).message,
		).toContain("spinCount must be an integer in 0..");
	});

	it("rejects a spinCount / workbookSpinCount above the uint32 ceiling (would emit invalid 1e+21 XML)", async () => {
		expect(
			(
				await writeErr({
					sheets: [
						{ name: "S", rows: [["a"]], protection: { sheet: true, spinCount: 1e21 } },
					],
				})
			).message,
		).toContain("spinCount must be an integer in 0..");
		expect(
			(
				await writeErr({
					sheets: [{ name: "S", rows: [["a"]] }],
					protection: { workbookSpinCount: 0x1_0000_0000 },
				})
			).message,
		).toContain("workbookSpinCount must be an integer in 0..");
	});

	it("rejects a bad workbook-protection flag and a non-boolean cell protection flag", async () => {
		expect(
			(
				await writeErr({
					sheets: [{ name: "S", rows: [["a"]] }],
					protection: { lockStructure: 1 },
				})
			).message,
		).toContain("protection.lockStructure must be a boolean");
		expect(
			(
				await writeErr({
					sheets: [
						{
							name: "S",
							rows: [[{ value: "x", style: { protection: { locked: 1 } } }]],
						},
					],
				})
			).message,
		).toContain("protection.locked must be a boolean");
	});
});
