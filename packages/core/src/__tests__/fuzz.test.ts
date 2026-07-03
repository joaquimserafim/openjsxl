import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError } from "../errors";
import { openXlsx, streamSheetRows } from "../reader/workbook";
import { createXmlStream, tokenize } from "../xml";
import { openZip } from "../zip";

// Robustness by fuzzing (F2.4e): for ANY input the tokenizer must not throw or hang, and the
// zip/reader must either parse or throw a typed XlsxError — never a bare TypeError/RangeError
// from an out-of-bounds read, and never crash or loop forever. Seeded so failures reproduce.

function rng(seed: number): () => number {
	let a = seed >>> 0;
	return () => {
		a = (a + 0x6d2b79f5) | 0;
		let t = Math.imul(a ^ (a >>> 15), 1 | a);
		t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

// Characters that stress the XML tokenizer's constructs (tags, entities, comments, CDATA, PIs).
const XML_CHARS = `<>/&;"'= \t\nabcxyz0![]CDATA-?xmlkr:`;

function randomXmlish(r: () => number, maxLen: number): string {
	let s = "";
	const len = Math.floor(r() * maxLen);
	for (let i = 0; i < len; i++) s += XML_CHARS.charAt(Math.floor(r() * XML_CHARS.length));
	return s;
}

function randomBytes(r: () => number, maxLen: number): Uint8Array {
	const len = Math.floor(r() * maxLen);
	const bytes = new Uint8Array(len);
	for (let i = 0; i < len; i++) bytes[i] = Math.floor(r() * 256);
	return bytes;
}

function mutate(base: Uint8Array, r: () => number): Uint8Array {
	const bytes = base.slice();
	const flips = 1 + Math.floor(r() * 8);
	for (let f = 0; f < flips; f++) bytes[Math.floor(r() * bytes.length)] = Math.floor(r() * 256);
	return bytes;
}

describe("fuzz — tokenizer never throws or hangs", () => {
	it("tokenizes arbitrary XML-ish strings to completion", () => {
		const r = rng(1);
		for (let n = 0; n < 3000; n++) {
			const s = randomXmlish(r, 80);
			expect(() => [...tokenize(s)]).not.toThrow();
		}
	});

	it("feeds arbitrary strings through the chunk stream in random splits", () => {
		const r = rng(2);
		for (let n = 0; n < 2000; n++) {
			const s = randomXmlish(r, 80);
			expect(() => {
				const stream = createXmlStream();
				for (let i = 0; i < s.length; ) {
					const size = 1 + Math.floor(r() * 6);
					stream.push(s.slice(i, i + size));
					i += size;
				}
				stream.flush();
			}).not.toThrow();
		}
	});
});

describe("fuzz — openZip only ever throws XlsxError", () => {
	it("on random byte buffers", () => {
		const r = rng(3);
		for (let n = 0; n < 3000; n++) {
			const bytes = randomBytes(r, 260);
			try {
				openZip(bytes);
			} catch (e) {
				expect(e).toBeInstanceOf(XlsxError);
			}
		}
	});

	it("on a mutated valid archive (including reading each part)", async () => {
		const r = rng(4);
		const base = await loadFixture("basic.xlsx");
		for (let n = 0; n < 400; n++) {
			try {
				const zip = openZip(mutate(base, r));
				for (const name of zip.entries.keys()) {
					try {
						await zip.read(name);
					} catch (e) {
						expect(e).toBeInstanceOf(XlsxError);
					}
				}
			} catch (e) {
				expect(e).toBeInstanceOf(XlsxError);
			}
		}
	});
});

describe("fuzz — openXlsx only ever throws XlsxError", () => {
	it("on a mutated valid workbook (draining every row)", async () => {
		const r = rng(5);
		const base = await loadFixture("basic.xlsx");
		for (let n = 0; n < 400; n++) {
			try {
				const wb = await openXlsx(mutate(base, r));
				// Drain every sheet fully — decoding each row/cell must not throw on garbage, and
				// a corrupted row count must not loop forever (a hang shows up as a test timeout).
				for (const info of wb.sheets) {
					let seen = 0;
					for await (const _row of wb.sheet(info.name).rows()) seen++;
					expect(seen).toBeGreaterThanOrEqual(0);
				}
			} catch (e) {
				expect(e).toBeInstanceOf(XlsxError);
			}
		}
	});
});

describe("regression — adversarial inputs the random seed cannot reach", () => {
	// Caught by an adversarial review, not the seeded fuzz: a column ref long enough to overflow
	// columnToIndex to a non-integer used to be *returned* (not thrown), poisoning lastCol; a
	// following cell without an `r` attribute then formatted { col: Infinity }, throwing a bare
	// Error out of the read path. The reader must degrade gracefully — no throw, or an XlsxError.
	// edge-overflow-col.xlsx is a valid package whose one row's first cell uses a 300-letter
	// column ref, followed by a cell with no `r`.
	it("reads the row via positional fallback instead of throwing (Worksheet.rows())", async () => {
		const bytes = await loadFixture("edge-overflow-col.xlsx");
		const wb = await openXlsx(bytes);
		const rows = [];
		for await (const row of wb.sheet("S").rows()) rows.push(row);
		// One row survives; the overflowing ref is kept verbatim and the next cell is positioned.
		expect(rows).toHaveLength(1);
		expect(rows[0]?.cells).toHaveLength(2);
	});

	it("reads the row via positional fallback instead of throwing (streamSheetRows())", async () => {
		const bytes = await loadFixture("edge-overflow-col.xlsx");
		const rows = [];
		for await (const row of streamSheetRows(bytes)) rows.push(row);
		expect(rows).toHaveLength(1);
		expect(rows[0]?.cells).toHaveLength(2);
	});
});

describe("fuzz — streamSheetRows only ever throws XlsxError", () => {
	// The constant-memory streaming reader (F2.2) is a separate path from Worksheet.rows(): it
	// tokenizes the worksheet chunk by chunk rather than as one string. Fuzz it directly so the
	// streaming row/cell state machine is exercised on garbage, not just the in-memory path.
	it("on a mutated valid workbook (draining the stream)", async () => {
		const r = rng(6);
		const base = await loadFixture("basic.xlsx");
		for (let n = 0; n < 400; n++) {
			try {
				let seen = 0;
				for await (const _row of streamSheetRows(mutate(base, r))) seen++;
				expect(seen).toBeGreaterThanOrEqual(0);
			} catch (e) {
				expect(e).toBeInstanceOf(XlsxError);
			}
		}
	});
});
