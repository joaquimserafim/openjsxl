// 14 — Workbook fidelity (1.0): defined names, autofilters, protection, and print setup — author them,
// read them back, and round-trip them. Each is the SAME record the reader returns and the writer
// accepts, so read → modify → write carries them across untouched.
//
//   node 14-names-autofilter-protection-print.mjs   (from ./examples)
//   pnpm --filter openjsxl-examples m10

import { openXlsx, workbookToInput, writeXlsx } from "openjsxl";

// (1) Author a report sheet: a defined name, filter dropdowns, cell + sheet protection, and print setup.
const bytes = await writeXlsx({
	sheets: [
		{
			name: "Q3",
			rows: [
				["Region", "Revenue", "Notes"],
				["North", 1200, ""],
				["South", 900, ""],
				["East", 1500, ""],
			],
			// Filter dropdowns over the table (F10.2). The hidden _xlnm._FilterDatabase Excel pairs with
			// it is synthesized for you — you only supply the range.
			autoFilter: { ref: "A1:C4" },
			// Sheet protection (F10.3): lock the sheet. Cells are locked by default; to leave a cell
			// editable under protection, give it a style with `protection: { locked: false }`. (A real
			// password sets the hash attributes, which are carried verbatim — never computed.)
			protection: { sheet: true, formatCells: false },
			// Print setup (F10.4): landscape, fit to one page wide, with a header and footer.
			pageSetup: { orientation: "landscape", fitToWidth: 1, fitToHeight: 0 },
			pageMargins: {
				left: 0.5,
				right: 0.5,
				top: 0.75,
				bottom: 0.75,
				header: 0.3,
				footer: 0.3,
			},
			printOptions: { horizontalCentered: true, gridLines: true },
			headerFooter: { oddHeader: "&CQ3 Revenue", oddFooter: "&RPage &P of &N" },
		},
	],
	// Workbook-level defined name (F10.1): name a range so formulas can reference `RevenueRange`.
	definedNames: [{ name: "RevenueRange", refersTo: "Q3!$B$2:$B$4" }],
	// Workbook protection (F10.3): stop structure changes (adding/removing/reordering sheets).
	protection: { lockStructure: true },
});

// (2) Read it all back — every record comes back in the same shape it went in.
const wb = await openXlsx(bytes);
const sheet = wb.sheet("Q3");

console.log("defined names:  ", wb.definedNames.map((n) => `${n.name} → ${n.refersTo}`).join(", "));
console.log("workbook lock:  ", JSON.stringify(wb.protection));
console.log("autofilter:     ", JSON.stringify(sheet.autoFilter));
console.log("sheet protection:", JSON.stringify(sheet.protection));
console.log("page setup:     ", JSON.stringify(sheet.pageSetup));
console.log("header/footer:  ", JSON.stringify(sheet.headerFooter));

// (3) Round-trip through the bridge: read → workbookToInput → write keeps all of it.
const round = await openXlsx(await writeXlsx(await workbookToInput(wb)));
const same =
	round.definedNames.length === 1 &&
	round.sheet("Q3").autoFilter?.ref === "A1:C4" &&
	round.sheet("Q3").pageSetup?.orientation === "landscape";
console.log("\nround-trips losslessly:", same);

// (4) Macro-enabled sources (F10.5): openjsxl reads .xlsm but writes only .xlsx, so a rewrite drops the
// VBA macros. Check `Workbook.macroEnabled` to warn before converting.
if (wb.macroEnabled) {
	console.warn("this workbook has VBA macros — rewriting to .xlsx will drop them");
} else {
	console.log("macroEnabled:   ", wb.macroEnabled, "(a plain .xlsx — nothing to warn about)");
}
