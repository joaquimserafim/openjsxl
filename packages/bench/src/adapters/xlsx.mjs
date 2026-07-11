import * as XLSX from "xlsx";

// SheetJS adapter (xlsx@0.18.5 — the last version published to the public npm registry). Reading
// materializes every row via sheet_to_json; writing deflates (compression:true) so its output size
// is comparable. Cell styling is a SheetJS Pro feature absent from this build, so styled write is
// declared unsupported rather than measured doing less work than the others.

export const capabilities = { read: true, write: true, writeStyled: false, stream: false };

// SheetJS sniffs the container (xlsx / xlsb / ods / csv) from the bytes, so ONE read path serves
// every format lane — the `format` argument is accepted for a uniform adapter signature but unused.
export async function read(bytes, _format = "xlsx") {
	const wb = XLSX.read(bytes, { type: "buffer" });
	let sink = 0;
	for (const name of wb.SheetNames) {
		const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true });
		for (const row of rows) for (const v of row) sink = mix(sink, v);
	}
	return sink;
}

export async function write(dataset, kind) {
	if (kind === "styled") {
		throw new Error("unsupported: SheetJS community build does not emit cell styles");
	}
	const ws = XLSX.utils.aoa_to_sheet(dataset);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "Bench");
	const out = XLSX.write(wb, { type: "buffer", bookType: "xlsx", compression: true });
	return new Uint8Array(out);
}

function mix(acc, v) {
	if (typeof v === "number") return acc + v;
	if (typeof v === "string") return acc + v.length;
	return v ? acc + 1 : acc;
}
