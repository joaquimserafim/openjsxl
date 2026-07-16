import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openCsv } from "../csv";
import { openXlsx } from "../workbook";
import { parseSheetProtection } from "../worksheet";

// F10.3 — read <sheetProtection>, <workbookProtection>, and xf <protection>; carry password material
// verbatim; degrade for non-xlsx formats.

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const sheet = (body: string): string =>
	`<?xml version="1.0"?><worksheet ${NS}><sheetData/>${body}</worksheet>`;

describe("parseSheetProtection", () => {
	it("reads boolean flags (1/0) and password material verbatim", () => {
		const body =
			'<sheetProtection sheet="1" formatCells="0" sort="1" selectLockedCells="0" ' +
			'password="C258" algorithmName="SHA-512" hashValue="h==" saltValue="s==" spinCount="100000"/>';
		expect(parseSheetProtection(sheet(body))).toEqual({
			sheet: true,
			formatCells: false,
			sort: true,
			selectLockedCells: false,
			password: "C258",
			algorithmName: "SHA-512",
			hashValue: "h==",
			saltValue: "s==",
			spinCount: 100000,
		});
	});

	it("returns undefined when there is no <sheetProtection>", () => {
		expect(parseSheetProtection(sheet(""))).toBeUndefined();
	});

	// spinCount is xsd:unsignedInt; a hostile 21-digit value must be DROPPED, not parsed to a float that
	// would re-emit as `1e+21` (adversarial-review regression). A valid uint32 max is kept.
	it("drops an out-of-range spinCount but keeps a valid uint32", () => {
		expect(
			parseSheetProtection(
				sheet('<sheetProtection sheet="1" spinCount="999999999999999999999"/>'),
			),
		).toEqual({ sheet: true });
		expect(
			parseSheetProtection(sheet('<sheetProtection sheet="1" spinCount="4294967295"/>')),
		).toEqual({ sheet: true, spinCount: 4294967295 });
	});

	// A <customSheetView> can also carry a <sheetProtection>-like nesting; only a direct <worksheet>
	// child counts (the parseAutoFilter/parseFreezePane scoping precedent).
	it("ignores a <sheetProtection> nested inside <customSheetViews>", () => {
		const nested =
			'<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}">' +
			'<sheetProtection sheet="1"/></customSheetView></customSheetViews>';
		expect(parseSheetProtection(sheet(nested))).toBeUndefined();
	});
});

describe("reader — openpyxl-authored protected fixture (F10.3)", () => {
	it("reads sheet + workbook protection and per-cell locked/hidden; password carried verbatim", async () => {
		const wb = await openXlsx(await loadFixture("openpyxl-protection.xlsx"));
		const ws = wb.sheet("Data");
		expect(ws.protection?.sheet).toBe(true);
		expect(ws.protection?.formatCells).toBe(false);
		expect(ws.protection?.password).toBe("C258"); // openpyxl's hash of "hunter2", kept verbatim
		expect(wb.protection).toEqual({ lockStructure: true });
		// A2 was explicitly unlocked, A3 hidden.
		expect(ws.style("A2")?.protection).toEqual({ locked: false, hidden: false });
		expect(ws.style("A3")?.protection).toEqual({ locked: true, hidden: true });
	});
});

describe("reader — degrade for non-xlsx formats", () => {
	it("openCsv exposes sheet + workbook protection as undefined", () => {
		const wb = openCsv(new TextEncoder().encode("a,b\n1,2\n"));
		expect(wb.protection).toBeUndefined();
		expect(wb.sheet(wb.sheets[0]?.name ?? "").protection).toBeUndefined();
	});
});
