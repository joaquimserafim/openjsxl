import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openCsv } from "../csv";
import { openXlsx } from "../workbook";
import { parseAutoFilter } from "../worksheet";

// F10.2 — sheet-level autoFilter: read the filter range, drop a hostile/non-canonical ref, and strip the
// paired _xlnm._FilterDatabase from Workbook.definedNames (it is surfaced as Worksheet.autoFilter instead).

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet ${NS}><sheetData/>${body}</worksheet>`;

describe("parseAutoFilter", () => {
	it("reads a canonical range", () => {
		expect(parseAutoFilter(sheet('<autoFilter ref="A1:C10"/>'))).toEqual({ ref: "A1:C10" });
	});

	it("reads a single-cell range", () => {
		expect(parseAutoFilter(sheet('<autoFilter ref="B2"/>'))).toEqual({ ref: "B2" });
	});

	it("keeps the range but ignores filterColumn/sortState children (criteria are a documented drop)", () => {
		const body =
			'<autoFilter ref="A1:C4"><filterColumn colId="2"><filters><filter val="NY"/></filters>' +
			'</filterColumn><sortState ref="A1:C4"><sortCondition ref="B2:B4"/></sortState></autoFilter>';
		expect(parseAutoFilter(sheet(body))).toEqual({ ref: "A1:C4" });
	});

	it("returns undefined when there is no autoFilter", () => {
		expect(parseAutoFilter(sheet(""))).toBeUndefined();
	});

	it("DROPS a hostile / non-canonical / out-of-grid ref (never returns something the writer rejects)", () => {
		expect(parseAutoFilter(sheet('<autoFilter ref="a1:c10"/>'))).toBeUndefined(); // lowercase
		expect(parseAutoFilter(sheet('<autoFilter ref="A1:ZZZZ9"/>'))).toBeUndefined(); // past XFD
		expect(parseAutoFilter(sheet('<autoFilter ref="A1:B1048577"/>'))).toBeUndefined(); // past last row
		expect(parseAutoFilter(sheet("<autoFilter/>"))).toBeUndefined(); // no ref
		expect(parseAutoFilter(sheet('<autoFilter ref=""/>'))).toBeUndefined(); // empty ref
	});

	it("DROPS a backwards range (bottom-right before top-left) — the writer rejects it, so the shared bound holds", () => {
		// Canonical + in-grid but reversed. Before this drop the reader surfaced it and the strict
		// writer then rejected the very value the reader returned, breaking read → bridge → write.
		expect(parseAutoFilter(sheet('<autoFilter ref="B2:A1"/>'))).toBeUndefined();
		expect(parseAutoFilter(sheet('<autoFilter ref="C10:A1"/>'))).toBeUndefined(); // both axes reversed
		expect(parseAutoFilter(sheet('<autoFilter ref="A2:A1"/>'))).toBeUndefined(); // only the row reversed
	});

	// Adversarial-review regression (F10.2): <autoFilter> is ALSO a legal child of <customSheetView>, so a
	// flat scan would fabricate a sheet-level filter a saved custom view merely retained. Only a DIRECT
	// child of <worksheet> counts.
	it("ignores an <autoFilter> nested inside <customSheetViews> (a saved view, not the active filter)", () => {
		const nested =
			'<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}" ' +
			'filter="1" showAutoFilter="1"><autoFilter ref="A1:C3"/></customSheetView></customSheetViews>';
		expect(parseAutoFilter(sheet(nested))).toBeUndefined();
	});

	it("still reads the sheet-level filter when a custom view ALSO carries one (depth-1 wins)", () => {
		const body =
			'<autoFilter ref="A1:B9"/>' +
			'<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}">' +
			'<autoFilter ref="Z1:Z9"/></customSheetView></customSheetViews>';
		expect(parseAutoFilter(sheet(body))).toEqual({ ref: "A1:B9" });
	});
});

describe("reader — openpyxl-authored filtered fixture (F10.2)", () => {
	it("surfaces the filter range, drops criteria/sort, and strips _xlnm._FilterDatabase", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-autofilter.xlsx"));
		// The fixture carries <filterColumn> criteria + <sortState>; only the range survives.
		expect(wb.sheet("Data").autoFilter).toEqual({ ref: "A1:C4" });
		// openpyxl stored _xlnm._FilterDatabase in workbook.xml; it must NOT leak into definedNames.
		expect(wb.definedNames).toEqual([]);
	});
});

describe("reader — degrade for non-xlsx formats", () => {
	it("openCsv exposes autoFilter as undefined", () => {
		const wb = openCsv(new TextEncoder().encode("a,b\n1,2\n"));
		expect(wb.sheet(wb.sheets[0]?.name ?? "").autoFilter).toBeUndefined();
	});
});
