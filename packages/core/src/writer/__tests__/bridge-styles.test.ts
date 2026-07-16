import { readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { MAX_TABLE_NAME_LEN, tableNameProblem } from "../../ooxml/table";
import { openCsv } from "../../reader/csv";
import { openOds } from "../../reader/ods";
import { openXlsx, type Workbook } from "../../reader/workbook";
import { openXlsb } from "../../reader/xlsb";
import { uniquifyTableName, workbookToInput } from "../from-workbook";
import { writeXlsx } from "../workbook";

// F4.4 — the bridge carries styles. Contract: read → workbookToInput → writeXlsx → read gives a
// deep-equal style(ref) for every populated cell (values/types were already lossless since F3.3),
// and an UNSTYLED workbook still rewrites to byte-identical archives. The openpyxl-authored
// fixture is the acid test: real-producer styles, theme+tint colors, custom number formats.

// Pictures are compared by anchor + mime + name + a content DIGEST, never the raw bytes: a real
// image is megabytes, and the digest (length + FNV-1a) proves the exact bytes survived without
// bloating the snapshot. Distinct bytes hash apart; identical bytes (deduped media) hash the same.
function imageDigest(bytes: Uint8Array): string {
	let h = 0x811c9dc5;
	for (let i = 0; i < bytes.length; i++) {
		h ^= bytes[i] as number;
		h = Math.imul(h, 0x01000193);
	}
	return `${bytes.length}:${(h >>> 0).toString(16)}`;
}

async function styleSnapshot(wb: Workbook) {
	const out: Record<string, Record<string, unknown>> = {};
	for (const info of wb.sheets) {
		const sheet = wb.sheet(info.name);
		const cells: Record<string, unknown> = {};
		for await (const row of sheet.rows()) {
			for (const cell of row.cells) {
				cells[cell.ref] = {
					// Error cells write as their literal text (documented F3.3 flattening), so the
					// comparable identity of an 'error' cell IS the string it becomes.
					type: cell.type === "error" ? "string" : cell.type,
					value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
					style: sheet.style(cell.ref),
					// Formula TEXT (F5.4) is part of the fidelity contract too.
					formula: sheet.formula(cell.ref),
				};
			}
		}
		out[info.name] = {
			cells,
			// Geometry (F4.5) is part of the fidelity contract too.
			columns: sheet.columns,
			rowProperties: Object.fromEntries(sheet.rowProperties),
			freeze: sheet.freeze,
			// Structural metadata (F4.6): merges, hyperlinks, and tab visibility carry across.
			mergedCells: sheet.mergedCells,
			hyperlinks: sheet.hyperlinks,
			state: info.state,
			// Comments (F5.2) round-trip too — ref, resolved author, and plain text.
			comments: sheet.comments,
			// Tables (F9.1): name, ref, columns, header/totals flags, and style carry across.
			tables: sheet.tables,
			// Data validations (F9.2): sqref, type, operands, prompts/errors carry across.
			dataValidations: sheet.dataValidations,
			// Conditional formatting (F9.3): sqref, rules, inline dxf carry across.
			conditionalFormatting: sheet.conditionalFormatting,
			// autoFilter (F10.2): the filter range round-trips (criteria/sort are a documented drop).
			autoFilter: sheet.autoFilter,
			// Pictures (F6.4): anchor + mime + name + a content digest (not raw bytes).
			images: (await sheet.images()).map((img) => ({
				anchor: img.anchor,
				mime: img.mime,
				name: img.name,
				digest: imageDigest(img.bytes),
			})),
		};
	}
	return out;
}

async function rewrite(wb: Workbook): Promise<Uint8Array> {
	return writeXlsx(await workbookToInput(wb));
}

// A VALUE-and-structure snapshot for the cross-format conversion property (F7.4). Unlike
// styleSnapshot, it deliberately omits styles/number-formats/geometry: the ods/xlsb/csv readers
// don't carry those, and the writer adds an implicit date format on the way out, so comparing them
// across a conversion would diverge for reasons that aren't data loss. The contract for a non-xlsx
// source is exactly this: values, types, merges, hyperlinks, and tab state survive the trip to .xlsx.
async function dataSnapshot(wb: Workbook) {
	const out: Record<string, unknown> = {};
	for (const info of wb.sheets) {
		const sheet = wb.sheet(info.name);
		const cells: Record<string, unknown> = {};
		for await (const row of sheet.rows()) {
			for (const cell of row.cells) {
				cells[cell.ref] = {
					// Error cells flatten to their literal text on write (documented F3.3), so an
					// 'error' cell's comparable identity IS the string it becomes.
					type: cell.type === "error" ? "string" : cell.type,
					value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
				};
			}
		}
		out[info.name] = {
			cells,
			mergedCells: sheet.mergedCells,
			hyperlinks: sheet.hyperlinks,
			state: info.state,
		};
	}
	return out;
}

describe("bridge — styles round-trip", () => {
	it("carries every style of the openpyxl-authored fixture (acid test)", async () => {
		const before = await openXlsx(await loadFixture("openpyxl-styled.xlsx"));
		const snap = await styleSnapshot(before);
		const after = await openXlsx(await rewrite(before));
		expect(await styleSnapshot(after)).toEqual(snap);

		// Spot-check the hard cases survived: theme+tint, custom numFmt, full-load cell.
		const sheet = after.sheet("Styled");
		expect(sheet.style("B2")?.font?.color).toEqual({ theme: 4, tint: 0.3999755851924192 });
		expect(sheet.style("C4")?.numberFormat).toBe('"kg" 0.0');
		expect(sheet.style("C5")?.fill).toEqual({
			patternType: "solid",
			fgColor: { rgb: "FFDDEBF7" },
		});
	});

	it("carries a styled BLANK cell across the bridge", async () => {
		const style = {
			border: { top: { style: "medium" } },
			fill: { patternType: "gray125" },
		} as const;
		const first = await writeXlsx({
			sheets: [{ name: "S", rows: [["a", { value: null, style }]] }],
		});
		const again = await openXlsx(await rewrite(await openXlsx(first)));
		expect(again.sheet("S").cell("B1").type).toBe("empty");
		expect(again.sheet("S").style("B1")).toEqual(style);
	});

	it("flattens row/column DEFAULT styles into per-cell styles (documented)", async () => {
		// col-row-styles.xlsx styles bare cells via <col style> and <row s customFormat> defaults.
		// The bridge writes each cell's EFFECTIVE style directly; the rewritten file has no
		// defaults but every cell reads back with the same format as before.
		const before = await openXlsx(await loadFixture("col-row-styles.xlsx"));
		const after = await openXlsx(await rewrite(before));
		const sheet = after.sheet("Sheet1");
		expect(sheet.numberFormat("B1")).toBe("mm-dd-yy"); // was column-default
		expect(sheet.cell("B1").type).toBe("date");
		expect(sheet.numberFormat("A3")).toBe("0.00%"); // was row-default
		expect(sheet.style("A1")).toBeUndefined(); // unstyled stays unstyled
	});

	it("rewrites an UNSTYLED workbook to byte-identical archives (with and without dates)", async () => {
		// The implicit date format round-trips through style() as {numberFormat:'mm-dd-yy'}, which
		// reverse-maps to the same built-in id 14 — so even date-bearing bare input reproduces the
		// exact bytes, not merely equivalent ones.
		for (const input of [
			{ sheets: [{ name: "S", rows: [["a", 1, true], [3.14]] }] },
			{ sheets: [{ name: "S", rows: [["x", new Date(Date.UTC(2020, 0, 1))]] }] },
		]) {
			const first = await writeXlsx(input);
			const second = await rewrite(await openXlsx(first));
			expect(Array.from(second)).toEqual(Array.from(first));
		}
	});
});

describe("bridge — pictures round-trip (F6.4)", () => {
	it("carries anchored pictures across read → bridge → write", async () => {
		const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
		const jpg = new Uint8Array([0xff, 0xd8, 0xff, 9, 9]);
		const first = await writeXlsx({
			sheets: [
				{
					name: "Pics",
					rows: [["hi"]],
					images: [
						{
							anchor: { from: { col: 2, row: 3 }, ext: { cx: 100, cy: 200 } },
							bytes: png,
							mime: "image/png",
							name: "Logo",
						},
						{
							anchor: { from: { col: 4, row: 4 }, to: { col: 6, row: 8 } },
							bytes: jpg,
							mime: "image/jpeg",
						},
					],
				},
			],
		});
		const before = await openXlsx(first);
		const after = await openXlsx(await rewrite(before));
		// The reader shape IS the writer input, so the pictures deep-equal across the round trip.
		expect(await after.sheet("Pics").images()).toEqual(await before.sheet("Pics").images());
		expect((await after.sheet("Pics").images()).length).toBe(2);
	});

	it("leaves an imageless workbook on the byte-identity path (images key never attached)", async () => {
		const first = await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1, true]] }] });
		const input = await workbookToInput(await openXlsx(first));
		expect(input.sheets.every((s) => !("images" in s))).toBe(true);
		expect(Array.from(await writeXlsx(input))).toEqual(Array.from(first));
	});
});

describe("bridge — corpus property", () => {
	it("every readable fixture round-trips losslessly OR fails typed — never bare, never silent", async () => {
		const dataDir = fileURLToPath(new URL("../../../../fixtures/data/", import.meta.url));
		const files = (await readdir(dataDir)).filter((f) => f.endsWith(".xlsx"));
		expect(files.length).toBeGreaterThan(5);
		let lossless = 0;
		const typedFailures: string[] = [];
		for (const file of files) {
			let before: Workbook;
			try {
				before = await openXlsx(await loadFixture(file));
			} catch {
				continue; // the intentionally-broken fixtures
			}
			let bytes: Uint8Array;
			try {
				bytes = await rewrite(before);
			} catch (e) {
				// A tolerated read the writer can't represent must surface as a TYPED error.
				expect(e, file).toBeInstanceOf(XlsxError);
				expect((e as XlsxError).code, file).toBe("invalid-input");
				typedFailures.push(file);
				continue;
			}
			const after = await openXlsx(bytes);
			expect(await styleSnapshot(after), file).toEqual(await styleSnapshot(before));
			lossless++;
		}
		expect(lossless).toBeGreaterThan(5);
		// The only fixture allowed to refuse: the fuzz file whose cell ref is a 300-letter column
		// (kept faithfully by the tolerant reader, but addressable nowhere on a writable grid).
		expect(typedFailures).toEqual(["edge-overflow-col.xlsx"]);
	});

	it("every readable ods/xlsb/csv fixture converts to xlsx losslessly or fails typed (F7.4)", async () => {
		// The corpus property, extended across formats: a non-xlsx file read → bridge → writeXlsx →
		// read must preserve values/types/merges/hyperlinks/state, or refuse with a TYPED error. The
		// intentionally-broken fixtures (encrypted / wrong-mimetype / no-content) throw on open and are
		// skipped, exactly like the broken .xlsx fixtures above.
		const dataDir = fileURLToPath(new URL("../../../../fixtures/data/", import.meta.url));
		const files = (await readdir(dataDir)).filter((f) => /\.(ods|xlsb|csv)$/.test(f));
		expect(files.length).toBeGreaterThan(3);
		const lossless: string[] = [];
		const typedFailures: string[] = [];
		for (const file of files) {
			const bytes = await loadFixture(file);
			let before: Workbook;
			try {
				before = file.endsWith(".csv")
					? openCsv(bytes)
					: file.endsWith(".ods")
						? await openOds(bytes)
						: await openXlsb(bytes);
			} catch {
				continue; // the intentionally-broken (typed-reject) fixtures throw on open
			}
			const snap = await dataSnapshot(before);
			let out: Uint8Array;
			try {
				out = await rewrite(before);
			} catch (e) {
				expect(e, file).toBeInstanceOf(XlsxError);
				expect((e as XlsxError).code, file).toBe("invalid-input");
				typedFailures.push(file);
				continue;
			}
			expect(await dataSnapshot(await openXlsx(out)), file).toEqual(snap);
			lossless.push(file);
		}
		// Every readable ods/xlsb/csv fixture converts losslessly — NONE is allowed to typed-reject on
		// write. Pinning the (empty) reject set catches a lossless→reject regression that a bare count
		// would hide (the F7.4-review hole). The key fixtures must be in the lossless set, not silently
		// skipped as if they had thrown on open.
		expect(typedFailures).toEqual([]);
		for (const key of [
			"equiv.ods",
			"equiv.xlsb",
			"equiv.csv",
			"xlsb-basic.xlsb",
			"ods-edge.ods",
			"basic.csv",
		]) {
			expect(lossless, key).toContain(key);
		}
	});
});

describe("bridge — hostile files (review regressions)", () => {
	// Hand-craft a minimal workbook whose sheet1.xml is `sheetXml`, through the writer's own zip
	// layer — for inputs the tolerant reader accepts but no writer should ever have produced.
	async function craftWorkbook(sheetXml: string): Promise<Uint8Array> {
		const { writeZip } = await import("../zip");
		const enc = new TextEncoder();
		const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		return writeZip([
			{
				name: "[Content_Types].xml",
				data: enc.encode(
					`${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
				),
			},
			{
				name: "_rels/.rels",
				data: enc.encode(
					`${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/workbook.xml",
				data: enc.encode(
					`${decl}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
				),
			},
			{
				name: "xl/_rels/workbook.xml.rels",
				data: enc.encode(
					`${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
				),
			},
			{
				name: "xl/worksheets/sheet1.xml",
				data: enc.encode(
					`${decl}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetXml}</sheetData></worksheet>`,
				),
			},
		]);
	}

	async function bridgeError(sheetXml: string): Promise<XlsxError> {
		const wb = await openXlsx(await craftWorkbook(sheetXml));
		const err = await workbookToInput(wb).then(
			() => null,
			(e) => e,
		);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("invalid-input");
		return err as XlsxError;
	}

	it("refuses a cell beyond Excel's grid quickly and typed, instead of hanging", async () => {
		// A ref like A99999999999999 PARSES fine — the danger is that its row number becomes the
		// length of the rows array, which the writer then iterates.
		const started = Date.now();
		const err = await bridgeError(
			'<row r="99999999999999"><c r="A99999999999999"><v>1</v></c></row>',
		);
		expect(Date.now() - started).toBeLessThan(2000); // typed refusal, not an hours-long loop
		expect(err.message).toMatch(/grid position/);
	});

	it("refuses case-variant duplicate refs instead of silently dropping one value", async () => {
		// "A1" and "a1" are DISTINCT cells to the reader (cell() keys by the verbatim ref) but one
		// grid slot to a writer — last-wins placement would make the value 2 vanish with no error.
		const err = await bridgeError(
			'<row r="1"><c r="A1"><v>2</v></c><c r="a1"><v>1</v></c></row>',
		);
		expect(err.message).toMatch(/"A1" and "a1".*one grid position/);
	});

	it("keeps last-wins for SAME-spelling duplicate refs, matching the reader's cell()", async () => {
		const wb = await openXlsx(
			await craftWorkbook('<row r="1"><c r="A1"><v>2</v></c><c r="A1"><v>1</v></c></row>'),
		);
		expect(wb.sheet("S").cell("A1").value).toBe(1); // reader's own answer is last-wins
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(wb)));
		expect(rewritten.sheet("S").cell("A1").value).toBe(1); // bridge agrees
	});
});

describe("bridge — F9.5 table round-trip hardening", () => {
	it("normalizes a foreign table's illegal name so it re-saves (openpyxl cell-ref name 'A1')", async () => {
		const before = await openXlsx(await loadFixture("openpyxl-table-oddname.xlsx"));
		const table = before.sheet("Sheet1").tables[0];
		// The reader normalized the cell-ref-shaped displayName "A1" into a legal identifier ON READ —
		// so what it returns is already something the strict writer accepts.
		expect(table?.name).not.toBe("A1");
		expect(tableNameProblem(table?.name ?? "")).toBeUndefined();
		expect(table?.ref).toBe("A1:C4");
		expect(table?.columns.map((c) => c.name)).toEqual(["Item", "Qty", "City"]);
		// The point of F9.5: the whole file now RE-SAVES (this threw invalid-input before) and re-reads,
		// and the normalized name is a stable fixpoint.
		const after = await openXlsx(await rewrite(before));
		const t2 = after.sheet("Sheet1").tables[0];
		expect(t2?.name).toBe(table?.name);
		expect(t2?.ref).toBe("A1:C4");
		expect(t2?.columns.map((c) => c.name)).toEqual(["Item", "Qty", "City"]);
	});

	// The reader's name normalization (F9.5) can map two DISTINCT illegal names to the SAME legal string
	// (e.g. "" and control-chars both → "_"); the bridge dedupes so the writer doesn't reject the
	// duplicate (review HIGH). A unique name passes through unchanged (byte-identity).
	it("dedupes table names that collide after normalization (bridge, workbook-wide)", () => {
		const seen = new Set<string>();
		expect(uniquifyTableName("Sales", seen)).toBe("Sales"); // first: unchanged
		expect(uniquifyTableName("_", seen)).toBe("_"); // first collision-prone name: unchanged
		expect(uniquifyTableName("_", seen)).toBe("__2"); // second "_" → suffixed
		expect(uniquifyTableName("_", seen)).toBe("__3"); // third → next suffix
		expect(uniquifyTableName("sales", seen)).toBe("sales_2"); // case-insensitive collision with "Sales"
		expect(tableNameProblem("__2")).toBeUndefined(); // suffixed names stay legal
		expect(tableNameProblem("sales_2")).toBeUndefined();
	});

	// Truncating a near-max name to make room for the suffix slices at UTF-16 code units — a cut
	// landing inside an astral character used to leave a lone high surrogate, an ILLEGAL candidate the
	// writer then rejects (F9.6 regression). The half pair is dropped; the candidate is always legal.
	it("keeps a suffixed candidate legal when the length cut lands inside an astral pair", () => {
		// 2 + 126×2 + 1 = 255 units (the max); the "_2" suffix cut at 253 units lands mid-pair.
		const name = `AB${"😀".repeat(126)}C`;
		expect(name.length).toBe(MAX_TABLE_NAME_LEN);
		expect(tableNameProblem(name)).toBeUndefined(); // the input itself is legal
		const seen = new Set<string>([name.toLowerCase()]);
		const candidate = uniquifyTableName(name, seen);
		expect(candidate.endsWith("_2")).toBe(true);
		expect(candidate.length).toBeLessThanOrEqual(MAX_TABLE_NAME_LEN);
		expect(tableNameProblem(candidate)).toBeUndefined(); // no lone surrogate — writer-legal
	});
});
