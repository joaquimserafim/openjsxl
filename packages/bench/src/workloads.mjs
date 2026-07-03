// Workload definitions + deterministic data generators, shared by every library adapter so all
// of them serialize (and read back) the EXACT same logical data. No randomness — a workload is a
// pure function of (kind, rows, cols) so runs reproduce and every library is handed identical work.

// Fixed 10 columns; the three sizes are 10k / 100k / 1M CELLS.
export const COLS = 10;
export const SIZES = [
	{ key: "10k", cells: 10_000, rows: 1_000 },
	{ key: "100k", cells: 100_000, rows: 10_000 },
	{ key: "1M", cells: 1_000_000, rows: 100_000 },
];

export const WORKLOADS = ["numbers", "strings", "styled"];

// A tiny word pool → strings with realistic repetition, so a shared-strings table (Excel, ExcelJS,
// SheetJS) is meaningfully exercised rather than being handed 1M unique strings or 1M copies of one.
const WORDS = [
	"Alpha",
	"Bravo",
	"Charlie",
	"Delta",
	"Echo",
	"Foxtrot",
	"Golf",
	"Hotel",
	"India",
	"Juliet",
	"Kilo",
	"Lima",
	"Mike",
	"November",
	"Oscar",
	"Papa",
	"Quebec",
	"Romeo",
	"Sierra",
	"Tango",
	"Uniform",
	"Victor",
	"Whiskey",
	"Xray",
	"Yankee",
	"Zulu",
];

// Three canonical styles cycled by column, so every writer's style interner sees a handful of
// distinct-but-repeated styles (the realistic case) rather than one style or a million. Shape is
// openjsxl's CellStyle; the ExcelJS adapter translates it (see adapters/exceljs.mjs).
export const STYLES = [
	{
		font: { bold: true, color: { rgb: "FF1F4E79" } },
		fill: { patternType: "solid", fgColor: { rgb: "FFDDEBF7" } },
	},
	{ font: { italic: true }, numberFormat: "#,##0.00" },
	{
		font: { bold: true },
		fill: { patternType: "solid", fgColor: { rgb: "FFFCE4D6" } },
		alignment: { horizontal: "right" },
	},
];

// One cell's logical value, purely a function of its coordinates.
function numberAt(r, c) {
	// A float with two decimals, spread across a wide range — no trivially-compressible constant.
	return Math.round((r * COLS + c) * 150 + c * 7 + 0.25) / 100;
}
function stringAt(r, c) {
	return `${WORDS[(r * COLS + c) % WORDS.length]}-${r % 500}`;
}

// Produce ONE row (0-based row index r) for the given workload, as the array a writer consumes.
export function rowAt(kind, r) {
	const row = new Array(COLS);
	for (let c = 0; c < COLS; c++) {
		if (kind === "numbers") row[c] = numberAt(r, c);
		else if (kind === "strings") row[c] = stringAt(r, c);
		else row[c] = { value: numberAt(r, c), style: STYLES[c % STYLES.length] };
	}
	return row;
}

/**
 * Materialize the whole dataset as a rows array — what a buffered writer (writeXlsx, ExcelJS,
 * SheetJS) is handed. Built OUTSIDE the timed region so generation never counts against a library.
 */
export function buildDataset(kind, rows) {
	const out = new Array(rows);
	for (let r = 0; r < rows; r++) out[r] = rowAt(kind, r);
	return out;
}

/**
 * A LAZY row source (generator) for the streaming writer — the honest streaming use case, where
 * rows arrive from a cursor and the full array never exists. Each row is created on demand.
 */
export function* datasetGenerator(kind, rows) {
	for (let r = 0; r < rows; r++) yield rowAt(kind, r);
}
