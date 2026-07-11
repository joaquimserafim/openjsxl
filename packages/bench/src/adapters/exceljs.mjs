import ExcelJS from "exceljs";

// ExcelJS adapter (v4.4.0). The standard in-memory Workbook API — what most ExcelJS users write —
// so the comparison reflects real usage, not its lower-level streaming writer.

export const capabilities = { read: true, write: true, writeStyled: true, stream: false };

// ExcelJS reads xlsx and csv only — it has no xlsb/ods parser, so those lanes decline (thrown →
// reported as n/a, never measured doing less work).
export async function read(bytes, format = "xlsx") {
	const wb = new ExcelJS.Workbook();
	if (format === "csv") {
		const { Readable } = await import("node:stream");
		await wb.csv.read(Readable.from([Buffer.from(bytes)]));
	} else if (format === "xlsx") {
		await wb.xlsx.load(bytes);
	} else {
		throw new Error(`unsupported: ExcelJS cannot read ${format}`);
	}
	let sink = 0;
	wb.eachSheet((ws) => {
		ws.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				sink = mix(sink, cell.value);
			});
		});
	});
	return sink;
}

export async function write(dataset, kind) {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet("Bench");
	if (kind !== "styled") {
		for (const row of dataset) ws.addRow(row);
	} else {
		for (let r = 0; r < dataset.length; r++) {
			const src = dataset[r];
			const row = ws.getRow(r + 1);
			for (let c = 0; c < src.length; c++) {
				const cell = row.getCell(c + 1);
				cell.value = src[c].value;
				applyStyle(cell, src[c].style);
			}
		}
	}
	const buf = await wb.xlsx.writeBuffer();
	return new Uint8Array(buf);
}

// Translate an openjsxl CellStyle into ExcelJS's style shape so both writers do equivalent work.
function applyStyle(cell, style) {
	if (!style) return;
	if (style.font) {
		const f = {};
		if (style.font.bold) f.bold = true;
		if (style.font.italic) f.italic = true;
		if (style.font.color?.rgb) f.color = { argb: style.font.color.rgb };
		cell.font = f;
	}
	if (style.fill?.patternType === "solid" && style.fill.fgColor?.rgb) {
		cell.fill = {
			type: "pattern",
			pattern: "solid",
			fgColor: { argb: style.fill.fgColor.rgb },
		};
	}
	if (style.numberFormat) cell.numFmt = style.numberFormat;
	if (style.alignment?.horizontal) cell.alignment = { horizontal: style.alignment.horizontal };
}

function mix(acc, v) {
	if (typeof v === "number") return acc + v;
	if (typeof v === "string") return acc + v.length;
	return v ? acc + 1 : acc;
}
