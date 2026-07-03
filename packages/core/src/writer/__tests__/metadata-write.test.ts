import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { writeXlsx } from "../workbook";

// F4.6 — structural metadata write: merged cells, hyperlinks (the writer's first per-sheet rels
// part), and tab visibility. Everything written must re-read through the reader's own accessors
// verbatim and carry across the bridge; a metadata-free sheet must keep its exact pre-F4.6 bytes
// (covered by the golden pins, re-asserted via the no-op normalization test here).

const decoder = new TextDecoder();

describe("writeXlsx — merged cells", () => {
	it("writes merges that re-read verbatim, in document order", async () => {
		const merges = ["A1:B2", "C1:C3", "A4:D4"] as const;
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "M", rows: [["a"]], merges }] }),
		);
		expect(wb.sheet("M").mergedCells).toEqual(merges);
	});

	it("emits <mergeCells> with a count, directly after </sheetData>", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "M", rows: [["a", "b"]], merges: ["A1:B1"] }],
		});
		const xml = decoder.decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"));
		expect(xml).toContain(
			'</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>',
		);
	});

	it("rejects malformed, single-cell, reversed, and out-of-grid ranges", async () => {
		const cases: [string, RegExp][] = [
			["banana", /not a canonical A1 range/],
			["A1", /not a canonical A1 range/], // a bare cell is not a range
			["a1:b2", /not a canonical A1 range/], // lowercase is not the canonical spelling
			["A0:B2", /not a canonical A1 range/],
			["A1:A1", /merges a single cell/],
			["B2:A1", /top-left to bottom-right/],
			["A1:XFE1", /within Excel's grid/], // XFE = column 16385, one past XFD
			["A1:B1048577", /within Excel's grid/], // one row past the grid
		];
		for (const [merge, pattern] of cases) {
			const err = await writeXlsx({
				sheets: [{ name: "M", rows: [[1]], merges: [merge] }],
			}).then(
				() => undefined,
				(e) => e,
			);
			expect(err, merge).toBeInstanceOf(XlsxError);
			expect((err as XlsxError).code, merge).toBe("invalid-input");
			expect((err as XlsxError).message, merge).toMatch(pattern);
		}
	});

	it("rejects overlapping merges, naming both ranges (Excel repairs overlap)", async () => {
		const err = await writeXlsx({
			sheets: [{ name: "M", rows: [[1]], merges: ["A1:B2", "D1:E1", "B2:C3"] }],
		}).then(
			() => undefined,
			(e) => e,
		);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).message).toMatch(/"A1:B2" and "B2:C3" overlap/);
	});

	it("rejects a duplicate range (a merge overlaps itself)", async () => {
		await expect(
			writeXlsx({ sheets: [{ name: "M", rows: [[1]], merges: ["A1:B2", "A1:B2"] }] }),
		).rejects.toMatchObject({ code: "invalid-input", message: /overlap/ });
	});

	it("detects overlap fast on adversarial volume (the sweep, not O(n²))", async () => {
		// 200k single-column vertical merges sharing rows — pairwise column-disjoint until the
		// last one collides. A naive pairwise check would grind; the sweep throws quickly.
		const merges: string[] = [];
		for (let i = 0; i < 200_000; i++) {
			const row = i * 4 + 1;
			merges.push(`A${row}:B${row + 1}`);
		}
		merges.push("A1:A2"); // collides with the very first range
		const started = Date.now();
		await expect(
			writeXlsx({ sheets: [{ name: "M", rows: [[1]], merges }] }),
		).rejects.toMatchObject({ code: "invalid-input", message: /overlap/ });
		expect(Date.now() - started).toBeLessThan(5000);
	});
});

describe("writeXlsx — hyperlinks", () => {
	it("writes external, in-workbook, and combined links that re-read verbatim", async () => {
		const hyperlinks = [
			{ ref: "A1", target: "https://example.com/?a=1&b=2", tooltip: "docs" },
			{ ref: "B2:C2", location: "'L'!A1", display: "jump" },
			{ ref: "D4", target: "mailto:x@example.com", location: "Sheet9!B5" },
		] as const;
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "L", rows: [["a"]], hyperlinks }] }),
		);
		expect(wb.sheet("L").hyperlinks).toEqual(hyperlinks);
	});

	it("emits the per-sheet rels part with TargetMode external, ids matching the sheet", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "L",
					rows: [["a"]],
					hyperlinks: [
						{ ref: "A1", target: "https://example.com/" },
						{ ref: "B1", location: "'L'!A1" }, // no rel — location only
						{ ref: "C1", target: 'https://e.com/?q="x"&y=1' },
					],
				},
			],
		});
		const zip = openZip(bytes);
		const sheet = decoder.decode(await zip.read("xl/worksheets/sheet1.xml"));
		expect(sheet).toContain(
			'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
		);
		expect(sheet).toContain(
			'<hyperlinks><hyperlink ref="A1" r:id="rId1"/>' +
				`<hyperlink ref="B1" location="'L'!A1"/>` +
				'<hyperlink ref="C1" r:id="rId2"/></hyperlinks>',
		);
		const rels = decoder.decode(await zip.read("xl/worksheets/_rels/sheet1.xml.rels"));
		expect(rels).toContain(
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
		);
		expect(rels).toContain(
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://e.com/?q=&quot;x&quot;&amp;y=1" TargetMode="External"/>',
		);
	});

	it("location-only links produce no rels part and no xmlns:r declaration", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "L", rows: [["a"]], hyperlinks: [{ ref: "A1", location: "B5" }] }],
		});
		const zip = openZip(bytes);
		expect(zip.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(false);
		const sheet = decoder.decode(await zip.read("xl/worksheets/sheet1.xml"));
		expect(sheet).not.toContain("xmlns:r");
	});

	it("carries an empty tooltip/display verbatim, and a newline in a tooltip survives", async () => {
		// The reader keeps tooltip:"" and display:"" (only location gates the empty string), and
		// literal newlines in attributes are emitted as &#10; so attribute-value normalization
		// can't turn them into spaces.
		const hyperlinks = [
			{ ref: "A1", target: "https://e.com/", tooltip: "", display: "" },
			{ ref: "B1", target: "https://e.com/", tooltip: "line1\nline2" },
		] as const;
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "L", rows: [["a"]], hyperlinks }] }),
		);
		expect(wb.sheet("L").hyperlinks).toEqual(hyperlinks);
	});

	it("rejects a link with no destination — including empty strings, which normalize away", async () => {
		for (const link of [
			{ ref: "A1" },
			{ ref: "A1", target: "" },
			{ ref: "A1", target: "", location: "" },
		]) {
			await expect(
				writeXlsx({ sheets: [{ name: "L", rows: [[1]], hyperlinks: [link] }] }),
			).rejects.toMatchObject({
				code: "invalid-input",
				message: /needs a target .* and\/or a location/,
			});
		}
	});

	it("rejects bad refs, unknown keys, non-strings, and XML-unsafe values", async () => {
		const cases: [object, RegExp][] = [
			[{ ref: "banana", target: "x" }, /not a canonical A1 cell or range/],
			[{ ref: "B2:A1", target: "x" }, /top-left to bottom-right/],
			[{ ref: "XFE1", target: "x" }, /within Excel's grid/],
			[{ target: "x" }, /ref must be a string/],
			[{ ref: "A1", target: "x", url: "y" }, /unknown property "url"/],
			[{ ref: "A1", target: "x", tooltip: 5 }, /tooltip must be a string/],
			[{ ref: "A1", target: "a\u0001b" }, /target contains a character not allowed in XML/],
		];
		for (const [link, pattern] of cases) {
			const err = await writeXlsx({
				sheets: [{ name: "L", rows: [[1]], hyperlinks: [link] } as never],
			}).then(
				() => undefined,
				(e) => e,
			);
			expect(err, JSON.stringify(link)).toBeInstanceOf(XlsxError);
			expect((err as XlsxError).message, JSON.stringify(link)).toMatch(pattern);
		}
	});
});

describe("writeXlsx — sheet visibility", () => {
	it("writes hidden and veryHidden states that re-read through info and accessor", async () => {
		const wb = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "Shown", rows: [[1]] },
					{ name: "H", rows: [], state: "hidden" },
					{ name: "VH", rows: [], state: "veryHidden" },
				],
			}),
		);
		expect(wb.sheets.map((s) => s.state)).toEqual(["visible", "hidden", "veryHidden"]);
		expect(wb.sheets.map((s) => s.visible)).toEqual([true, false, false]);
		expect(wb.sheet("VH").state).toBe("veryHidden");
	});

	it("emits state between sheetId and r:id, only for hiding states", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{ name: "A", rows: [], state: "visible" },
				{ name: "B", rows: [], state: "hidden" },
			],
		});
		const xml = decoder.decode(await openZip(bytes).read("xl/workbook.xml"));
		expect(xml).toContain('<sheet name="A" sheetId="1" r:id="rId1"/>');
		expect(xml).toContain('<sheet name="B" sheetId="2" state="hidden" r:id="rId2"/>');
	});

	it("aims the active tab at the first visible sheet when the first sheet is hidden", async () => {
		// The activeTab default is index 0 — pointing it at a hidden tab is what openpyxl also
		// avoids. No <bookViews> is emitted when the first sheet is visible.
		const firstHidden = await writeXlsx({
			sheets: [
				{ name: "H", rows: [], state: "hidden" },
				{ name: "V", rows: [[1]] },
			],
		});
		const xml = decoder.decode(await openZip(firstHidden).read("xl/workbook.xml"));
		expect(xml).toContain('<bookViews><workbookView activeTab="1"/></bookViews><sheets>');

		const laterHidden = await writeXlsx({
			sheets: [
				{ name: "V", rows: [[1]] },
				{ name: "H", rows: [], state: "veryHidden" },
			],
		});
		expect(decoder.decode(await openZip(laterHidden).read("xl/workbook.xml"))).not.toContain(
			"bookViews",
		);
	});

	it("rejects an unknown state and an all-hidden workbook", async () => {
		await expect(
			writeXlsx({ sheets: [{ name: "A", rows: [], state: "invisible" } as never] }),
		).rejects.toMatchObject({ code: "invalid-input", message: /state must be/ });
		await expect(
			writeXlsx({
				sheets: [
					{ name: "A", rows: [], state: "hidden" },
					{ name: "B", rows: [], state: "veryHidden" },
				],
			}),
		).rejects.toMatchObject({ code: "invalid-input", message: /at least one sheet/ });
	});

	it("reads state exactly ONCE — a value-flipping getter cannot dodge validation (F4.6 review)", async () => {
		// TOCTOU: a getter answering "visible" to validation and "hidden" to emission would write
		// the all-hidden workbook the guard exists to reject; one answering a validated literal
		// first and markup later would inject attributes (state is emitted unescaped BECAUSE it
		// can only be one of the three validated literals). Single-read makes both impossible.
		let reads = 0;
		const sheet = { name: "S", rows: [[1]] };
		Object.defineProperty(sheet, "state", {
			enumerable: true,
			get() {
				reads++;
				return reads === 1 ? "visible" : 'hidden" injected="1';
			},
		});
		const bytes = await writeXlsx({ sheets: [sheet as never] });
		expect(reads).toBe(1); // the single read saw "visible" — later values never exist
		const xml = decoder.decode(await openZip(bytes).read("xl/workbook.xml"));
		expect(xml).toContain('<sheet name="S" sheetId="1" r:id="rId1"/>');
		expect(xml).not.toContain("injected");
	});
});

describe("writeXlsx — metadata normalization and bridge", () => {
	it("normalizes no-op metadata away — the file matches a metadata-free write exactly", async () => {
		const plain = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] });
		const noop = await writeXlsx({
			sheets: [{ name: "S", rows: [[1]], merges: [], hyperlinks: [], state: "visible" }],
		});
		expect(Array.from(noop)).toEqual(Array.from(plain));
	});

	it("carries merges, hyperlinks, and visibility across the bridge", async () => {
		const first = await writeXlsx({
			sheets: [
				{
					name: "Main",
					rows: [["x", "y"]],
					merges: ["A2:B3"],
					hyperlinks: [
						{ ref: "A1", target: "https://example.com/", tooltip: "t" },
						{ ref: "B1", location: "'Hidden'!A1" },
					],
				},
				{ name: "Hidden", rows: [[1]], state: "hidden" },
			],
		});
		const again = await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(first))));
		const main = again.sheet("Main");
		expect(main.mergedCells).toEqual(["A2:B3"]);
		expect(main.hyperlinks).toEqual([
			{ ref: "A1", target: "https://example.com/", tooltip: "t" },
			{ ref: "B1", location: "'Hidden'!A1" },
		]);
		expect(again.sheets.map((s) => s.state)).toEqual(["visible", "hidden"]);
	});
});
