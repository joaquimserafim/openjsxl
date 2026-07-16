import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { openXlsx } from "../workbook";

// F10.5 — Workbook.macroEnabled: a read-only flag sniffed from the workbook content type, so a caller can
// warn before a rewrite discards VBA macros. openjsxl reads .xlsm but writes only .xlsx.

describe("Workbook.macroEnabled (F10.5)", () => {
	it("is true for a macro-enabled workbook (.xlsm content type)", async () => {
		const wb = await openXlsx(await loadFixture("crafted-macro-enabled.xlsm"));
		expect(wb.macroEnabled).toBe(true);
		// The workbook still opens and reads normally.
		expect(wb.sheet("Macros").cell("A1").value).toBe("hi");
	});

	it("is false for a plain .xlsx, and false after a rewrite (macros dropped, documented)", async () => {
		const plain = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] }));
		expect(plain.macroEnabled).toBe(false);

		// Rewriting an .xlsm through the bridge produces a plain .xlsx — the VBA project is dropped, and
		// the flag flips to false. This is documented, never a bare throw.
		const src = await openXlsx(await loadFixture("crafted-macro-enabled.xlsm"));
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(src)));
		expect(rewritten.macroEnabled).toBe(false);
		expect(rewritten.sheet("Macros").cell("A1").value).toBe("hi"); // data survives; only macros drop
	});
});
