// 13 — Tables, data validation & conditional formatting (0.9): author all three, read them back,
// and round-trip them. Each is the SAME record the reader returns and the writer accepts, so
// read → modify → write carries them across untouched.
//
//   node 13-tables-validation-formatting.mjs        (from ./examples)
//   pnpm --filter openjsxl-examples tables

import { openXlsx, workbookToInput, writeXlsx } from "openjsxl";

// (1) Author a sheet with a defined table, a dropdown, and a highlight rule.
const bytes = await writeXlsx({
	sheets: [
		{
			name: "Inventory",
			rows: [
				["Item", "Qty", "Status"],
				["Apples", 120, "In stock"],
				["Pears", 8, "Low"],
				["Plums", 0, "Out"],
			],
			// A defined table over the grid — column names DERIVE from the header row (`columns: []`).
			tables: [
				{
					name: "Stock",
					ref: "A1:C4",
					columns: [],
					headerRow: true,
					totalsRow: false,
					style: { name: "TableStyleMedium9", showRowStripes: true },
				},
			],
			// A dropdown on the Status column (an inline `list` — the quotes ARE part of the source text).
			dataValidations: [
				{
					sqref: ["C2:C4"],
					type: "list",
					formula1: '"In stock,Low,Out"',
					showDropDown: true,
				},
			],
			// Paint a low quantity (< 10) with a red fill (a `cellIs` highlight — its look is an inline dxf).
			conditionalFormatting: [
				{
					sqref: ["B2:B4"],
					rules: [
						{
							type: "cellIs",
							priority: 1,
							operator: "lessThan",
							formulas: ["10"],
							dxf: { fill: { bgColor: { rgb: "FFFFC7CE" } } },
						},
					],
				},
			],
		},
	],
});
console.log(`wrote ${bytes.length} bytes`);

// (2) Read them back — the accessors return exactly what we wrote.
const ws = (await openXlsx(bytes)).sheet("Inventory");
const table = ws.tables[0];
console.log(
	`table      : ${table.name} ${table.ref} [${table.columns.map((c) => c.name).join(", ")}]`,
);
const dv = ws.dataValidations[0];
console.log(`validation : ${dv.type} on ${dv.sqref.join(",")} = ${dv.formula1}`);
const rule = ws.conditionalFormatting[0].rules[0];
console.log(
	`cond. fmt  : ${rule.type} ${rule.operator} ${rule.formulas?.[0]} → fill ${rule.dxf?.fill?.bgColor?.rgb}`,
);

// (3) Round-trip: read → workbookToInput → write again. Tables/DV/CF all carry across the bridge.
const again = (await openXlsx(await writeXlsx(await workbookToInput(await openXlsx(bytes))))).sheet(
	"Inventory",
);
console.log(
	`round-trip : ${again.tables.length} table, ${again.dataValidations.length} validation, ${again.conditionalFormatting.length} conditional-format block — all preserved`,
);
