import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx } from "../workbook";
import { parseColumnProps, parseFreezePane, parseRowProperties } from "../worksheet";

// F4.5 — sheet geometry read: column widths/hidden, row heights/hidden, frozen panes. Unit tests
// drive the parsers with inline XML (degradation cases); the e2e block reads a real-producer
// fixture (openpyxl 3.1.5 — see fixtures/data/README.md).

const sheet = (inner: string): string =>
	`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${inner}</worksheet>`;

describe("parseColumnProps", () => {
	it("reads width, hidden, and both", () => {
		expect(
			parseColumnProps(
				sheet(
					'<cols><col min="2" max="3" width="25.5" customWidth="1"/>' +
						'<col min="5" max="5" hidden="1"/>' +
						'<col min="7" max="8" width="9.75" hidden="true"/></cols><sheetData/>',
				),
			),
		).toEqual([
			{ min: 2, max: 3, width: 25.5 },
			{ min: 5, max: 5, hidden: true },
			{ min: 7, max: 8, width: 9.75, hidden: true },
		]);
	});

	it("omits style-only entries and degrades out-of-bounds values", () => {
		expect(
			parseColumnProps(
				sheet(
					'<cols><col min="1" max="1" style="3"/>' + // style plumbing, not geometry
						'<col min="2" max="2" width="300"/>' + // width past Excel's 255 ceiling
						'<col min="0" max="2" width="10"/>' + // min < 1
						'<col min="3" max="2" width="10"/>' + // max < min
						'<col min="2" max="99999" width="10"/>' + // past the grid
						'<col min="4" max="4" width="banana" hidden="1"/></cols><sheetData/>', // bad width, hidden survives
				),
			),
		).toEqual([{ min: 4, max: 4, hidden: true }]);
	});
});

describe("parseRowProperties", () => {
	it("reads heights and hidden flags, keyed by row (explicit or positional)", () => {
		const map = parseRowProperties(
			sheet(
				'<sheetData><row r="2" ht="33" customHeight="1"><c r="A2"/></row>' +
					'<row ht="12.75"/>' + // no r → positional: row 3
					'<row r="9" hidden="1"/></sheetData>',
			),
		);
		expect(Object.fromEntries(map)).toEqual({
			2: { height: 33 },
			3: { height: 12.75 },
			9: { hidden: true },
		});
	});

	it("degrades bad heights and ignores rows with no geometry", () => {
		const map = parseRowProperties(
			sheet(
				'<sheetData><row r="1"><c r="A1"/></row>' + // plain row: absent
					'<row r="2" ht="0"/>' + // non-positive height
					'<row r="3" ht="500"/>' + // past Excel's 409.5 ceiling
					'<row r="4" ht="410" hidden="1"/></sheetData>', // bad height but hidden survives
			),
		);
		expect(Object.fromEntries(map)).toEqual({ 4: { hidden: true } });
	});
});

describe("parseFreezePane", () => {
	it("reads frozen rows/cols and their combinations", () => {
		const pane = (attrs: string) =>
			parseFreezePane(
				sheet(
					`<sheetViews><sheetView><pane ${attrs}/></sheetView></sheetViews><sheetData/>`,
				),
			);
		expect(pane('xSplit="1" ySplit="2" state="frozen"')).toEqual({ rows: 2, cols: 1 });
		expect(pane('ySplit="1" state="frozen"')).toEqual({ rows: 1 });
		expect(pane('xSplit="3" state="frozen"')).toEqual({ cols: 3 });
	});

	it("reads split (non-frozen) panes and degenerate freezes as undefined", () => {
		const pane = (attrs: string) =>
			parseFreezePane(
				sheet(
					`<sheetViews><sheetView><pane ${attrs}/></sheetView></sheetViews><sheetData/>`,
				),
			);
		expect(pane('xSplit="2000" ySplit="1000" state="split"')).toBeUndefined();
		expect(pane('xSplit="2000" ySplit="1000" state="frozenSplit"')).toBeUndefined();
		expect(pane('state="frozen"')).toBeUndefined(); // nothing actually frozen
		expect(pane('xSplit="1.5" state="frozen"')).toBeUndefined(); // non-integer count
		expect(parseFreezePane(sheet("<sheetData/>"))).toBeUndefined(); // no pane at all
	});
});

describe("geometry — openpyxl-authored fixture (e2e)", () => {
	it("reads columns, row properties, and the frozen pane of a real file", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-geometry.xlsx"));
		const geo = wb.sheet("Geo");
		expect(geo.columns).toEqual([
			{ min: 2, max: 2, width: 25.5 },
			{ min: 3, max: 3, width: 13, hidden: true },
			{ min: 4, max: 4, width: 9.75, hidden: true },
		]);
		expect(Object.fromEntries(geo.rowProperties)).toEqual({
			2: { height: 33 },
			4: { hidden: true }, // a property-only row: no cells in row 4
		});
		expect(geo.freeze).toEqual({ rows: 2, cols: 1 });

		const plain = wb.sheet("Plain");
		expect(plain.columns).toEqual([]);
		expect(plain.rowProperties.size).toBe(0);
		expect(plain.freeze).toBeUndefined();
	});
});

describe("geometry — scan scoping (F4.5 review regressions)", () => {
	it("ignores the <pane> of a saved Custom View (it lives after sheetData)", () => {
		// A user froze a row in a Custom View, then unfroze the normal view: the default
		// sheetView has no pane, but customSheetViews keeps one. Reporting it would fabricate a
		// freeze the active view doesn't have — and the bridge would write it for real.
		const xml = sheet(
			'<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>' +
				'<customSheetViews><customSheetView guid="{1}"><pane ySplit="1" topLeftCell="A2" state="frozen"/></customSheetView></customSheetViews>',
		);
		expect(parseFreezePane(xml)).toBeUndefined();
	});

	it("ignores col-named elements after sheetData (extension lists etc.)", () => {
		const xml = sheet(
			'<cols><col min="1" max="1" width="10"/></cols><sheetData/>' +
				'<extLst><ext><col min="2" max="2" width="99"/></ext></extLst>',
		);
		expect(parseColumnProps(xml)).toEqual([{ min: 1, max: 1, width: 10 }]);
	});

	it("keys row properties with the row assembler's exact r parsing (parseInt + fallback)", () => {
		// "1e3" parses as 1 to the assembler (parseInt) but 1000 to Number() — the height must
		// stay with the row the CELLS land on, or the bridge migrates it to a phantom row.
		const map = parseRowProperties(
			sheet(
				'<sheetData><row r="1e3" ht="42" customHeight="1"><c><v>1</v></c></row>' +
					'<row ht="10"/></sheetData>', // r-less: positional AFTER the same fallback chain
			),
		);
		expect(Object.fromEntries(map)).toEqual({ 1: { height: 42 }, 2: { height: 10 } });
	});
});
