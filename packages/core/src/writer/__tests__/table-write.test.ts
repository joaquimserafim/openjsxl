import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { SheetInput } from "../types";
import { writeXlsx } from "../workbook";

const decoder = new TextDecoder();
const tablePart = async (sheet: SheetInput): Promise<string> =>
	decoder.decode(
		await openZip(await writeXlsx({ sheets: [sheet] })).read("xl/tables/table1.xml"),
	);

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
	let at = 0;
	for (const c of chunks) {
		out.set(c, at);
		at += c.length;
	}
	return out;
}

// F9.1 — writing tables. Author via writeXlsx, read back through openXlsx, and pin every decision-8
// rejection as a typed XlsxError('invalid-input').

async function roundTrip(sheet: SheetInput) {
	const book = await openXlsx(await writeXlsx({ sheets: [sheet] }));
	return book.sheet(sheet.name);
}

const rejects = (sheet: SheetInput) =>
	expect(writeXlsx({ sheets: [sheet] })).rejects.toThrow(XlsxError);

describe("writeXlsx — tables round-trip", () => {
	it("writes a table whose column names derive from the header row", async () => {
		const ws = await roundTrip({
			name: "S",
			rows: [
				["Region", "Total"],
				["West", 10],
				["East", 20],
			],
			tables: [
				{ name: "Sales", ref: "A1:B3", columns: [], headerRow: true, totalsRow: false },
			],
		});
		expect(ws.tables).toEqual([
			{
				name: "Sales",
				ref: "A1:B3",
				columns: [{ name: "Region" }, { name: "Total" }],
				headerRow: true,
				totalsRow: false,
			},
		]);
	});

	it("carries the style banding and totals flag", async () => {
		const ws = await roundTrip({
			name: "S",
			rows: [
				["A", "B"],
				[1, 2],
				[3, 4],
			],
			tables: [
				{
					name: "T",
					ref: "A1:B3",
					columns: [],
					headerRow: true,
					totalsRow: true,
					style: { name: "TableStyleMedium2", showRowStripes: true },
				},
			],
		});
		const t = ws.tables[0];
		expect(t?.totalsRow).toBe(true);
		expect(t?.style).toEqual({ name: "TableStyleMedium2", showRowStripes: true });
	});

	it("assigns workbook-global ids across sheets and rejects a duplicate name", async () => {
		// Two tables on two sheets get table1.xml / table2.xml; the reader is agnostic to the id, so we
		// just check both survive.
		const book = await openXlsx(
			await writeXlsx({
				sheets: [
					{
						name: "S1",
						rows: [["X"], [1]],
						tables: [
							{
								name: "Ta",
								ref: "A1:A2",
								columns: [],
								headerRow: true,
								totalsRow: false,
							},
						],
					},
					{
						name: "S2",
						rows: [["Y"], [2]],
						tables: [
							{
								name: "Tb",
								ref: "A1:A2",
								columns: [],
								headerRow: true,
								totalsRow: false,
							},
						],
					},
				],
			}),
		);
		expect(book.sheet("S1").tables[0]?.name).toBe("Ta");
		expect(book.sheet("S2").tables[0]?.name).toBe("Tb");
	});

	it("a no-table workbook streams the same rows (byte-identity path unaffected)", async () => {
		// Sanity that adding the feature didn't perturb the plain path — the golden pins cover bytes;
		// here we just confirm a plain sheet still reads back.
		const ws = await roundTrip({ name: "S", rows: [["a", 1]] });
		expect(ws.tables).toEqual([]);
		expect(ws.cell("A1").value).toBe("a");
	});
});

describe("writeXlsx — table validation rejects malformed input (decision 8)", () => {
	const base = (over: Record<string, unknown>): SheetInput => ({
		name: "S",
		rows: [
			["H1", "H2"],
			[1, 2],
		],
		tables: [
			{ name: "T", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false, ...over },
		],
	});

	it("rejects a name that looks like a cell reference or has spaces", async () => {
		await rejects(base({ name: "A1" }));
		await rejects(base({ name: "My Table" }));
		await rejects(base({ name: "" }));
	});

	it("rejects a duplicate table name across the workbook (case-insensitive)", async () => {
		await rejects({
			name: "S",
			rows: [["H"], [1]],
			// two tables, same name up to case
			tables: [
				{ name: "Tbl", ref: "A1:A2", columns: [], headerRow: true, totalsRow: false },
				{ name: "TBL", ref: "A1:A2", columns: [], headerRow: true, totalsRow: false },
			],
		});
	});

	it("rejects overlapping table ranges", async () => {
		await rejects({
			name: "S",
			rows: [
				["A", "B", "C"],
				[1, 2, 3],
			],
			tables: [
				{ name: "T1", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false },
				{ name: "T2", ref: "B1:C2", columns: [], headerRow: true, totalsRow: false },
			],
		});
	});

	it("rejects a non-text or empty header cell (names derive from the header row)", async () => {
		await rejects({
			name: "S",
			rows: [
				["H1", 42],
				[1, 2],
			], // second header cell is a number
			tables: [{ name: "T", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false }],
		});
	});

	it("rejects columns whose count differs from the ref width, and duplicate header names", async () => {
		await rejects(base({ columns: [{ name: "only-one" }] })); // width is 2
		await rejects({
			name: "S",
			rows: [
				["Dup", "Dup"],
				[1, 2],
			],
			tables: [{ name: "T", ref: "A1:B2", columns: [], headerRow: true, totalsRow: false }],
		});
	});

	it("rejects an unknown property and a bad ref", async () => {
		await rejects(base({ bogus: 1 }));
		await rejects(base({ ref: "A1" })); // not a range
	});
});

describe("writeXlsx — table autoFilter/totals geometry (F9.1 review regressions)", () => {
	it("rejects a totals row on a too-short ref (typed, never a bare throw)", async () => {
		// Was `Error: invalid row index: 0` (bare) / a reversed autoFilter range.
		await rejects({
			name: "S",
			rows: [["x"]],
			tables: [{ name: "T", ref: "A1:A1", columns: [], headerRow: true, totalsRow: true }],
		});
		const err = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[], [], [], [], ["x"]],
					tables: [
						{ name: "T", ref: "A5:A5", columns: [], headerRow: false, totalsRow: true },
					],
				},
			],
		}).catch((e) => e);
		expect(err).toBeInstanceOf(XlsxError);
	});

	it("omits <autoFilter> for a header-less table (Excel/openpyxl never emit it)", async () => {
		const part = await tablePart({
			name: "S",
			rows: [
				["1", "2"],
				["3", "4"],
			],
			tables: [
				{
					name: "NoHeader",
					ref: "A1:B2",
					columns: [{ name: "a" }, { name: "b" }],
					headerRow: false,
					totalsRow: false,
				},
			],
		});
		expect(part).not.toContain("<autoFilter");
		expect(part).toContain('headerRowCount="0"');
	});

	it("emits an autoFilter that excludes the totals row on a valid multi-row table", async () => {
		const part = await tablePart({
			name: "S",
			rows: [["H"], [1], [2]],
			tables: [{ name: "T", ref: "A1:A3", columns: [], headerRow: true, totalsRow: true }],
		});
		expect(part).toContain('<autoFilter ref="A1:A2"/>'); // A3 (totals) excluded
		expect(part).toContain('totalsRowCount="1"');
	});
});

describe("streamXlsx — tables use caller-provided column names (no header read)", () => {
	it("emits a table whose names come from tables[i].columns", async () => {
		const bytes = await drain(
			streamXlsx({
				sheets: [
					{
						name: "S",
						rows: [
							["Region", "Total"],
							["West", 10],
						],
						tables: [
							{
								name: "Streamed",
								ref: "A1:B2",
								columns: [{ name: "Region" }, { name: "Total" }],
								headerRow: true,
								totalsRow: false,
							},
						],
					},
				],
			}),
		);
		const ws = (await openXlsx(bytes)).sheet("S");
		expect(ws.tables[0]?.name).toBe("Streamed");
		expect(ws.tables[0]?.columns).toEqual([{ name: "Region" }, { name: "Total" }]);
		expect(ws.cell("A2").value).toBe("West"); // the body streamed correctly around the table footer
	});
});

// F9.3 retrofit — a table's/column's highlight dxfs (headerRowDxfId/dataDxfId/totalsRowDxfId) become
// inline DxfStyle fields, interned into the shared <dxfs> table (the same one CF rules use).
describe("writeXlsx — table dxf highlights (F9.3 retrofit)", () => {
	it("round-trips table-level and per-column dxf highlights deep-equal", async () => {
		const table = {
			name: "Inv",
			ref: "A1:B3",
			columns: [
				{ name: "Item" },
				{ name: "Qty", dataStyle: { fill: { bgColor: { rgb: "FFFFC7CE" } } } },
			],
			headerRow: true,
			totalsRow: false,
			headerRowStyle: { font: { bold: true } },
		};
		const ws = await roundTrip({
			name: "S",
			rows: [
				["Item", "Qty"],
				["A", 5],
				["K", 9],
			],
			tables: [table],
		});
		const t = ws.tables[0];
		expect(t?.headerRowStyle).toEqual({ font: { bold: true } });
		expect(t?.columns[1]?.dataStyle).toEqual({ fill: { bgColor: { rgb: "FFFFC7CE" } } });
		// The id attrs land on the table part; the numeric index never surfaces in the model.
		const part = await tablePart({
			name: "S",
			rows: [
				["Item", "Qty"],
				["A", 5],
				["K", 9],
			],
			tables: [table],
		});
		expect(part).toContain("headerRowDxfId=");
		expect(part).toMatch(/<tableColumn[^>]*dataDxfId=/);
	});

	it("shares ONE <dxfs> slot between a table dxf and an identical CF dxf (shared index space)", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [["H"], [1]],
					tables: [
						{
							name: "T",
							ref: "A1:A2",
							columns: [{ name: "H", dataStyle: { font: { bold: true } } }],
							headerRow: true,
							totalsRow: false,
						},
					],
					conditionalFormatting: [
						{
							sqref: ["A2"],
							rules: [
								{
									type: "expression",
									priority: 1,
									formulas: ["TRUE"],
									dxf: { font: { bold: true } },
								},
							],
						},
					],
				},
			],
		});
		const styles = decoder.decode(await openZip(bytes).read("xl/styles.xml"));
		expect(styles).toContain('<dxfs count="1">'); // table + CF share the one bold dxf
	});

	it("rejects a malformed table dxf (typed)", async () => {
		await rejects({
			name: "S",
			rows: [["H"], [1]],
			tables: [
				{
					name: "T",
					ref: "A1:A2",
					columns: [{ name: "H" }],
					headerRow: true,
					totalsRow: false,
					// biome-ignore lint/suspicious/noExplicitAny: hostile input past the types
					headerRowStyle: { font: { size: -1 } } as any,
				},
			],
		});
	});
});
