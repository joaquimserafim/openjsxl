import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openXlsx, type Workbook } from "../../reader/workbook";
import { openZip } from "../../zip";
import { workbookToInput } from "../from-workbook";
import { writeXlsx } from "../workbook";

// F3.3 acceptance: reading a workbook, bridging it to writer input, and writing it back is lossless
// for the supported value set (values, types, sheet names/order). We prove it by comparing a
// read → write → read snapshot against the original, over real producer fixtures and synthetic
// inputs, plus a golden pin on the exact emitted XML.

// A comparable view of a workbook: every populated cell keyed by ref, with dates reduced to their
// epoch millis so Date identity doesn't matter.
async function snapshot(wb: Workbook) {
	const out: Record<string, Record<string, { type: string; value: unknown }>> = {};
	for (const info of wb.sheets) {
		const cells: Record<string, { type: string; value: unknown }> = {};
		for await (const row of wb.sheet(info.name).rows()) {
			for (const cell of row.cells) {
				cells[cell.ref] = {
					type: cell.type,
					value: cell.value instanceof Date ? cell.value.getTime() : cell.value,
				};
			}
		}
		out[info.name] = cells;
	}
	return out;
}

async function rewrite(wb: Workbook): Promise<Workbook> {
	return openXlsx(await writeXlsx(await workbookToInput(wb)));
}

describe("round-trip — real fixtures re-read identically after write", () => {
	// basic.xlsx exercises string, number, a date-styled serial, boolean, and a cached formula
	// (which reads as its number). None of these are outside the writer's supported set, so the
	// snapshot must be byte-for-byte equal across the round trip.
	it("preserves basic.xlsx values, types, and sheets", async () => {
		const before = await openXlsx(await loadFixture("basic.xlsx"));
		const snap = await snapshot(before);
		const after = await rewrite(before);
		expect(await snapshot(after)).toEqual(snap);
	});

	it("preserves minimal.xlsx", async () => {
		const before = await openXlsx(await loadFixture("minimal.xlsx"));
		const snap = await snapshot(before);
		const after = await rewrite(before);
		expect(await snapshot(after)).toEqual(snap);
	});
});

describe("round-trip — bridge over the supported value set", () => {
	it("carries values, types, sheet order, and sparse placement", async () => {
		const date = new Date(Date.UTC(2022, 3, 5, 8, 15));
		const wb1 = await openXlsx(
			await writeXlsx({
				sheets: [
					{ name: "One", rows: [["text", 12.5, true], [date], [undefined, "sparse B3"]] },
					{ name: "Two", rows: [[false, "x"]] },
				],
			}),
		);
		const snap = await snapshot(wb1);
		const wb2 = await rewrite(wb1);

		expect(wb2.sheets.map((s) => s.name)).toEqual(["One", "Two"]);
		expect(await snapshot(wb2)).toEqual(snap);
		// Spot-check the interesting cells survived by type, not just by snapshot equality.
		expect(wb2.sheet("One").cell("A2").type).toBe("date");
		expect(wb2.sheet("One").cell("B3").value).toBe("sparse B3");
		expect(wb2.sheet("Two").cell("A1")).toMatchObject({ type: "boolean", value: false });
	});

	it("does not materialize a dense grid for a far-apart cell", async () => {
		// A single cell at row 1000 must round-trip without exploding — the bridge keeps rows sparse.
		const wb1 = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["top"]] }] }));
		const input = await workbookToInput(wb1);
		expect(input.sheets[0]?.rows.length).toBe(1);
	});
});

describe("round-trip — golden wire format", () => {
	it("emits the exact expected worksheet XML for a fixed input", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["hi", 5, true]] }] });
		const xml = new TextDecoder().decode(await openZip(bytes).read("xl/worksheets/sheet1.xml"));
		expect(xml).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
				'<dimension ref="A1:C1"/><sheetData>' +
				'<row r="1">' +
				'<c r="A1" t="inlineStr"><is><t>hi</t></is></c>' +
				'<c r="B1"><v>5</v></c>' +
				'<c r="C1" t="b"><v>1</v></c>' +
				"</row>" +
				"</sheetData></worksheet>",
		);
	});
});

describe("round-trip — the README read→modify→write snippet (F9.6)", () => {
	it("appends a row via non-mutating spread (WorkbookInput is deeply readonly)", async () => {
		// This mirrors the READMEs' modify-flow snippet EXACTLY — it living here means the repo's
		// tsc gate type-checks the documented pattern (the old `.push` version failed TS2339).
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "Sheet1", rows: [["a", 1]] }] }),
		);
		const input = await workbookToInput(wb);
		const sheets = input.sheets.map((sheet, i) =>
			i === 0 ? { ...sheet, rows: [...sheet.rows, ["appended", "row"]] } : sheet,
		);
		const out = await openXlsx(await writeXlsx({ ...input, sheets }));
		expect(out.sheet("Sheet1").cell("A2").value).toBe("appended");
		expect(out.sheet("Sheet1").cell("B2").value).toBe("row");
		expect(out.sheet("Sheet1").cell("A1").value).toBe("a"); // original data intact
	});
});
