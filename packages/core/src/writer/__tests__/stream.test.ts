import { describe, expect, it } from "vitest";
import { openXlsx, type Workbook } from "../../reader/workbook";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F5.1 — streamXlsx must be reader-EQUIVALENT to writeXlsx (byte-identity is explicitly not a goal;
// the streamed layout uses data descriptors and omits <dimension>). We assert equivalence by reading
// both outputs and comparing a full snapshot: values, types, styles, geometry, metadata, formulas.

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

async function snapshot(wb: Workbook): Promise<unknown> {
	const out: Record<string, unknown> = {};
	for (const info of wb.sheets) {
		const sheet = wb.sheet(info.name);
		const cells: Record<string, unknown> = {};
		for await (const row of sheet.rows()) {
			for (const cell of row.cells) {
				cells[cell.ref] = {
					type: cell.type,
					value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
					style: sheet.style(cell.ref),
					formula: sheet.formula(cell.ref),
				};
			}
		}
		out[info.name] = {
			cells,
			columns: sheet.columns,
			rowProperties: Object.fromEntries(sheet.rowProperties),
			freeze: sheet.freeze,
			mergedCells: sheet.mergedCells,
			hyperlinks: sheet.hyperlinks,
			comments: sheet.comments,
			state: info.state,
		};
	}
	return out;
}

const style = { font: { bold: true, color: { rgb: "FFFF0000" } } } as const;

const INPUTS: { name: string; input: WorkbookInput; opts?: { date1904?: boolean } }[] = [
	{
		name: "values + types",
		input: {
			sheets: [
				{
					name: "S",
					rows: [
						["a", 1, true],
						[3.14, null],
					],
				},
			],
		},
	},
	{
		name: "dates (1904)",
		input: { sheets: [{ name: "S", rows: [["d", new Date(Date.UTC(2020, 0, 1))]] }] },
		opts: { date1904: true },
	},
	{
		name: "styled + theme",
		input: {
			sheets: [
				{
					name: "S",
					rows: [
						[
							{ value: "x", style },
							{ value: 1, style: { font: { color: { theme: 4, tint: 0.4 } } } },
						],
					],
				},
			],
		},
	},
	{
		name: "geometry",
		input: {
			sheets: [
				{
					name: "G",
					rows: [["a"], [], [null, "b"]],
					columns: [{ min: 1, max: 1, width: 20 }],
					// Row 7 is PAST the last data row — a trailing property-only <row/> both writers emit.
					rowProperties: { 1: { height: 30 }, 7: { height: 44, hidden: true } },
					freeze: { rows: 1, cols: 1 },
				},
			],
		},
	},
	{
		name: "metadata + comments",
		input: {
			sheets: [
				{
					name: "M",
					rows: [["a", "b"]],
					merges: ["A1:B1"],
					hyperlinks: [{ ref: "A1", target: "https://example.com", tooltip: "t" }],
					comments: [{ ref: "A1", author: "Ada", text: "note" }],
				},
				{ name: "Hidden", rows: [["x"]], state: "hidden" },
			],
		},
	},
	{
		name: "formulas",
		input: { sheets: [{ name: "F", rows: [[{ formula: "B1*2", value: 84 }, 42]] }] },
	},
];

describe("streamXlsx — reader equivalence with writeXlsx", () => {
	it("produces an equivalent reader snapshot for every corpus-shaped input", async () => {
		for (const { name, input, opts } of INPUTS) {
			const buffered = await openXlsx(await writeXlsx(input, opts));
			const streamed = await openXlsx(await drain(streamXlsx(input, opts)));
			expect(await snapshot(streamed), name).toEqual(await snapshot(buffered));
		}
	});

	it("streams sync- and async-iterable rows equivalently to a materialized array", async () => {
		const rows = [["a", 1], [2, "b"], [{ value: 3, style }]];
		function* syncGen(): Generator<(typeof rows)[number]> {
			for (const r of rows) yield r;
		}
		async function* asyncGen(): AsyncGenerator<(typeof rows)[number]> {
			for (const r of rows) yield r;
		}
		const base = await snapshot(
			await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] })),
		);
		const sync = await snapshot(
			await openXlsx(await drain(streamXlsx({ sheets: [{ name: "S", rows: syncGen() }] }))),
		);
		const async = await snapshot(
			await openXlsx(await drain(streamXlsx({ sheets: [{ name: "S", rows: asyncGen() }] }))),
		);
		expect(sync).toEqual(base);
		expect(async).toEqual(base);
	});

	it("errors the stream (does not hang) on invalid input", async () => {
		const bad = streamXlsx({ sheets: [] });
		await expect(drain(bad)).rejects.toThrow(/at least one sheet/);
		const badRows = streamXlsx({ sheets: [{ name: "S", rows: [42 as never] }] });
		await expect(drain(badRows)).rejects.toThrow(/must be an array/);
	});

	it("closes the async row source on early cancel (review regression: no cursor leak)", async () => {
		let yielded = 0;
		let closed = false;
		async function* infinite(): AsyncGenerator<readonly number[]> {
			try {
				for (let i = 1; ; i++) {
					yielded++;
					yield [i]; // never ends on its own — a DB cursor stand-in
				}
			} finally {
				closed = true; // the source's finally: closing the cursor
			}
		}
		const reader = streamXlsx({ sheets: [{ name: "S", rows: infinite() }] }).getReader();
		// Pull until the source is actually being consumed (past the local header, into compression),
		// so there is really an open "cursor" to leak.
		while (yielded === 0) await reader.read();
		await reader.cancel(); // must tear down: run the source's finally + release the compressor
		expect(closed).toBe(true);
	});

	it("emits rowProperties past the last streamed row (M5-analysis regression: silent drop)", async () => {
		// The stream ends at row 1, but rows 5 and 9 carry height/hidden. The buffered writer flushes
		// them as trailing property-only <row/> elements; the streamed writer must too — same input,
		// same metadata out.
		const input = {
			sheets: [
				{
					name: "S",
					rows: [["a"]],
					rowProperties: {
						1: { height: 30 },
						9: { height: 12 },
						5: { height: 44, hidden: true },
					},
				},
			],
		};
		const buffered = await openXlsx(await writeXlsx(input));
		const streamed = await openXlsx(await drain(streamXlsx(input)));
		expect([...streamed.sheet("S").rowProperties]).toEqual([
			...buffered.sheet("S").rowProperties,
		]);
		expect(streamed.sheet("S").rowProperties.get(5)).toEqual({ height: 44, hidden: true });
		expect(streamed.sheet("S").rowProperties.get(9)).toEqual({ height: 12 });
	});

	it("emits the VALIDATED sheet name, not one a getter flips to afterwards (review regression: TOCTOU)", async () => {
		let reads = 0;
		const flip = {
			get name() {
				reads++;
				return reads === 1 ? "Good" : "Bad/Name"; // a forbidden name after validation
			},
			rows: [["x"]],
		};
		// Both writers must carry the name validated at read #1 ("Good"), never the flipped one.
		const streamed = await openXlsx(await drain(streamXlsx({ sheets: [flip] })));
		expect(streamed.sheets.map((s) => s.name)).toEqual(["Good"]);
		reads = 0;
		const buffered = await openXlsx(await writeXlsx({ sheets: [flip] }));
		expect(buffered.sheets.map((s) => s.name)).toEqual(["Good"]);
	});
});

// F6.1 — the per-sheet part/rel wiring is single-sourced (sheetRelPlumbing + sheetSideParts), so the
// buffered and streaming writers CANNOT drift on which OPC parts they emit or how a sheet's rels are
// ordered. These guard exactly that invariant: reader-equivalence (above) wouldn't catch a part name
// that changed in lockstep, and the body golden pins don't cover the side parts.
describe("per-sheet part wiring (F6.1 dedup guard)", () => {
	// One sheet exercises every side part (hyperlink + comments → rels/comments/vml), one is
	// hyperlinks-only (rels only), one is bare (no side parts).
	const wb: WorkbookInput = {
		sheets: [
			{
				name: "M",
				rows: [["a", "b"]],
				hyperlinks: [{ ref: "A1", target: "https://x.io" }],
				comments: [{ ref: "A1", author: "Z", text: "n" }],
			},
			{ name: "Links", rows: [["c"]], hyperlinks: [{ ref: "A1", target: "https://y.io" }] },
			{ name: "Plain", rows: [["d"]] },
		],
	};
	const partNames = (bytes: Uint8Array): string[] => [...openZip(bytes).entries.keys()].sort();

	it("both writers emit the identical set of OPC parts", async () => {
		const buffered = partNames(await writeXlsx(wb));
		const streamed = partNames(await drain(streamXlsx(wb)));
		expect(streamed).toEqual(buffered); // the drift guard: neither writer emits a part the other omits
		// …and that set is exactly the expected per-sheet wiring (names owned by sheetSideParts).
		expect(buffered).toEqual([
			"[Content_Types].xml",
			"_rels/.rels",
			"xl/_rels/workbook.xml.rels",
			"xl/comments1.xml",
			"xl/drawings/vmlDrawing1.vml",
			"xl/workbook.xml",
			"xl/worksheets/_rels/sheet1.xml.rels",
			"xl/worksheets/_rels/sheet2.xml.rels",
			"xl/worksheets/sheet1.xml",
			"xl/worksheets/sheet2.xml",
			"xl/worksheets/sheet3.xml",
		]);
	});

	it("sheet1 rels are ordered hyperlink (rId1), comments (rId2), vmlDrawing (rId3)", async () => {
		const bytes = await writeXlsx(wb);
		const rels = new TextDecoder().decode(
			await openZip(bytes).read("xl/worksheets/_rels/sheet1.xml.rels"),
		);
		const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		expect(rels).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
				`<Relationship Id="rId1" Type="${R}/hyperlink" Target="https://x.io" TargetMode="External"/>` +
				`<Relationship Id="rId2" Type="${R}/comments" Target="../comments1.xml"/>` +
				`<Relationship Id="rId3" Type="${R}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
				"</Relationships>",
		);
	});
});
