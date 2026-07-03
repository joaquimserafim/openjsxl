// 09 — Comments, formulas, and theme colors (0.5): the last fidelity gaps, closed.
//
//   node 09-comments-formulas-theme.mjs   (from ./examples)
//   pnpm --filter openjsxl-examples fidelity
//
// A cell can carry a formula (`{ formula, value? }` — the cached `value` is what readers that
// don't recalculate will see), a sheet takes `comments` (written together with the legacy VML
// drawing part Excel needs to actually SHOW them on hover), and `Workbook.resolveColor` turns a
// raw `{ theme, tint }` style color into the 8-digit ARGB Excel renders — resolved against the
// workbook's OWN theme, which the bridge also carries byte-identically through a rewrite.

import { openXlsx, workbookToInput, writeXlsx } from "openjsxl";

const bytes = await writeXlsx({
	sheets: [
		{
			name: "Invoice",
			rows: [
				[
					{ value: "Net", style: { font: { bold: true, color: { theme: 4 } } } },
					{ value: "VAT", style: { font: { bold: true } } },
					{ value: "Total", style: { font: { bold: true } } },
				],
				// B2 and C2 are live formulas; 23% VAT on A2. Cached values keep non-recalculating
				// consumers (and our own reader) seeing numbers immediately.
				[100, { formula: "A2*0.23", value: 23 }, { formula: "A2+B2", value: 123 }],
			],
			comments: [
				{ ref: "A2", author: "Ada", text: "net of discount" },
				{ ref: "C2", text: "auto-computed" }, // author is optional
			],
		},
	],
});

// Read it back — formulas, comments, and the theme all come from the file itself.
const wb = await openXlsx(bytes);
const sheet = wb.sheet("Invoice");

console.log("B2 formula :", sheet.formula("B2"), "→ cached", sheet.cell("B2").value);
console.log("C2 formula :", sheet.formula("C2"), "→ cached", sheet.cell("C2").value);
console.log("comments   :", JSON.stringify(sheet.comments));

// The style keeps the RAW theme reference (round-trip faithful); resolveColor gives the pixels.
const a1 = sheet.style("A1");
console.log("A1 color   :", JSON.stringify(a1.font.color), "→", wb.resolveColor(a1.font.color));
console.log("tinted 40% :", wb.resolveColor({ theme: 4, tint: 0.4 }));

// And the whole lot survives read → modify → write: the bridge carries formulas, comments, and
// the workbook's theme part verbatim.
const input = await workbookToInput(wb);
const again = await openXlsx(await writeXlsx(input));
console.log(
	"round-trip :",
	JSON.stringify(again.sheet("Invoice").formula("C2")),
	"+ comments:",
	again.sheet("Invoice").comments.length,
);
