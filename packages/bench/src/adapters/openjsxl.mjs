import { openCsv, openOds, openXlsb, openXlsx, streamXlsx, writeXlsx } from "openjsxl";

// openjsxl adapter — the library under test. Read materializes every cell value (forcing the full
// parse); buffered write returns the bytes; the streamed write drains a ReadableStream fed a LAZY
// row source, retaining no chunk, so its memory column reflects true constant-memory streaming.

export const capabilities = { read: true, write: true, writeStyled: true, stream: true };

// One explicit opener per container — the whole point of the multi-format read lanes is that each
// format has its own typed entry point (no auto-dispatch cost in the shipped core).
const OPENERS = { xlsx: openXlsx, xlsb: openXlsb, ods: openOds, csv: openCsv };

export async function read(bytes, format = "xlsx") {
	const open = OPENERS[format];
	if (!open) throw new Error(`unsupported: openjsxl has no opener for ${format}`);
	const wb = await open(bytes);
	let sink = 0;
	for (const info of wb.sheets) {
		const sheet = wb.sheet(info.name);
		for await (const row of sheet.rows()) {
			for (const cell of row.cells) sink = mix(sink, cell.value);
		}
	}
	return sink;
}

export async function write(dataset, _kind) {
	return await writeXlsx({ sheets: [{ name: "Bench", rows: dataset }] });
}

// `rows` is a fresh single-use iterable (a generator) — the array of rows never exists at once.
export async function writeStream(rows) {
	const stream = streamXlsx({ sheets: [{ name: "Bench", rows }] });
	const reader = stream.getReader();
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.length; // count bytes; drop the chunk immediately (constant memory)
	}
	return { length: total };
}

// Fold a value into a running checksum so the read work can't be optimized away.
function mix(acc, v) {
	if (typeof v === "number") return acc + v;
	if (typeof v === "string") return acc + v.length;
	return v ? acc + 1 : acc;
}
