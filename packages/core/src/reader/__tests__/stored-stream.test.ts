import { describe, expect, it } from "vitest";
import { crc32 } from "../../utils/crc32";
import { writeXlsx } from "../../writer";
import { writeZip, type ZipInput } from "../../writer/zip";
import { openZip } from "../../zip";
import { openXlsx, streamSheetRows } from "../workbook";

// F11.98: a stored (method 0) worksheet streams in 16 KiB slices, like a deflated one, and with F11.69
// a U+FEFF that starts a slice or an inflate chunk stays text.
// Python's zipfile stores parts unless told otherwise. Before the fix a stored sheet of a few thousand
// rows reached the XML stream in one chunk and failed with a false part-too-large.

const SHEET = "xl/worksheets/sheet1.xml";

/**
 * The parts of the archive `bytes`, uncompressed; a part named in `replace` holds that text instead.
 *
 * @example (await partsOf(xlsx, { [SHEET]: "<worksheet/>" })).find((p) => p.name === SHEET)?.data // the bytes of "<worksheet/>"
 */
async function partsOf(
	bytes: Uint8Array,
	replace: Readonly<Record<string, string>>,
): Promise<ZipInput[]> {
	const zip = openZip(bytes);
	const parts: ZipInput[] = [];
	for (const name of zip.entries.keys()) {
		const text = Object.hasOwn(replace, name) ? replace[name] : undefined;
		const data = text === undefined ? await zip.read(name) : new TextEncoder().encode(text);
		parts.push({ name, data });
	}
	return parts;
}

/**
 * A copy of the archive `bytes` with every part stored (method 0), as `zipfile.ZIP_STORED` writes it;
 * a part named in `replace` holds that text instead. The CRC is the production one: this file pins
 * streaming, and the zip suite pins the CRC against an independent table.
 *
 * @example (await storedCopy(xlsx)).byteLength > xlsx.byteLength // true: nothing is deflated
 */
async function storedCopy(
	bytes: Uint8Array,
	replace: Readonly<Record<string, string>> = {},
): Promise<Uint8Array> {
	const enc = new TextEncoder();
	const parts = (await partsOf(bytes, replace)).map((p) => ({
		name: enc.encode(p.name),
		data: p.data,
	}));
	const localSize = parts.reduce((n, p) => n + 30 + p.name.length + p.data.length, 0);
	const centralSize = parts.reduce((n, p) => n + 46 + p.name.length, 0);
	const out = new Uint8Array(localSize + centralSize + 22);
	const view = new DataView(out.buffer);
	let local = 0;
	let central = localSize;
	for (const { name, data } of parts) {
		const crc = crc32(data);
		view.setUint32(local, 0x04034b50, true);
		view.setUint16(local + 4, 20, true);
		view.setUint32(local + 14, crc, true);
		view.setUint32(local + 18, data.length, true);
		view.setUint32(local + 22, data.length, true);
		view.setUint16(local + 26, name.length, true);
		out.set(name, local + 30);
		out.set(data, local + 30 + name.length);
		view.setUint32(central, 0x02014b50, true);
		view.setUint16(central + 4, 20, true);
		view.setUint16(central + 6, 20, true);
		view.setUint32(central + 16, crc, true);
		view.setUint32(central + 20, data.length, true);
		view.setUint32(central + 24, data.length, true);
		view.setUint16(central + 28, name.length, true);
		view.setUint32(central + 42, local, true);
		out.set(name, central + 46);
		local += 30 + name.length + data.length;
		central += 46 + name.length;
	}
	view.setUint32(central, 0x06054b50, true);
	view.setUint16(central + 8, parts.length, true);
	view.setUint16(central + 10, parts.length, true);
	view.setUint32(central + 12, centralSize, true);
	view.setUint32(central + 16, localSize, true);
	return out;
}

/**
 * Where the chunks of `name`'s stream end inside the part, the last one excepted.
 *
 * @example await chunkEnds(stored, SHEET) // [16384, 32768, …] for a stored part over 32 KiB
 */
async function chunkEnds(bytes: Uint8Array, name: string): Promise<number[]> {
	const ends: number[] = [];
	let at = 0;
	for await (const chunk of openZip(bytes).readStream(name)) {
		at += chunk.byteLength;
		ends.push(at);
	}
	return ends.slice(0, -1);
}

async function buffered(bytes: Uint8Array) {
	const rows = [];
	for await (const row of (await openXlsx(bytes)).sheet("Data").rows()) rows.push(row);
	return rows;
}

async function streamed(bytes: Uint8Array) {
	const rows = [];
	for await (const row of streamSheetRows(bytes, "Data", { verifyChecksums: true }))
		rows.push(row);
	return rows;
}

describe("streamSheetRows on a stored worksheet (F11.98)", () => {
	it("streams 20,000 rows exactly as openXlsx reads them", async () => {
		const rows = Array.from({ length: 20_000 }, (_, r) => [
			r + 1,
			`row ${r + 1}`,
			r / 4,
			r % 2 === 0,
			-r,
		]);
		const bytes = await storedCopy(await writeXlsx({ sheets: [{ name: "Data", rows }] }));
		const sheet = openZip(bytes).entries.get(SHEET);
		expect(sheet?.method).toBe(0);
		// Over 1 MiB: HEAD failed from about 0.8 MB.
		expect(sheet?.compressedSize).toBeGreaterThan(1024 * 1024);

		const expected = await buffered(bytes);
		expect(expected).toHaveLength(20_000);
		expect(expected[19_999]?.cells.map((cell) => cell.value)).toEqual([
			20_000,
			"row 20000",
			19_999 / 4,
			false,
			-19_999,
		]);
		expect(await streamed(bytes)).toEqual(expected);
	});

	it("reads inline text whose UTF-8 bytes a slice boundary splits", async () => {
		let xml =
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
		for (let r = 1; r <= 20_000; r++) {
			xml += `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>漢字 ${r} é</t></is></c><c r="B${r}"><v>${r}</v></c></row>`;
		}
		xml += "</sheetData></worksheet>";
		const base = await writeXlsx({ sheets: [{ name: "Data", rows: [[1]] }] });
		const bytes = await storedCopy(base, { [SHEET]: xml });

		// At least one 16 KiB boundary falls inside a multi-byte character, so the case is exercised.
		const data = new TextEncoder().encode(xml);
		const split = [];
		for (let at = 16 * 1024; at < data.length; at += 16 * 1024) {
			if (((data[at] as number) & 0xc0) === 0x80) split.push(at);
		}
		expect(split.length).toBeGreaterThan(0);

		const expected = await buffered(bytes);
		expect(expected).toHaveLength(20_000);
		expect(expected[12_345]?.cells.map((cell) => cell.value)).toEqual(["漢字 12346 é", 12_346]);
		expect(await streamed(bytes)).toEqual(expected);
	});

	// U+FEFF runs, 40 at the start of every cell's text: about 230 KB, so a stored sheet has 14 slices.
	const BOM = "\uFEFF";
	let bomXml =
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
	for (let r = 1; r <= 1_200; r++) {
		bomXml += `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>${BOM.repeat(40)}${r}</t></is></c></row>`;
	}
	bomXml += "</sheetData></worksheet>";

	it.each([
		["stored", 0],
		["deflated", 8],
	])(
		"streams U+FEFF text across %s chunk boundaries as openXlsx reads it (F11.69)",
		async (_, method) => {
			const base = await writeXlsx({ sheets: [{ name: "Data", rows: [[1]] }] });
			const bytes =
				method === 0
					? await storedCopy(base, { [SHEET]: bomXml })
					: await writeZip(await partsOf(base, { [SHEET]: bomXml }));
			expect(openZip(bytes).entries.get(SHEET)?.method).toBe(method);

			// U+FEFF is the only non-ASCII character here, so a chunk boundary followed by a byte >= 0x80
			// falls in a U+FEFF run: the next chunk's text then begins with U+FEFF, which HEAD dropped.
			const data = new TextEncoder().encode(bomXml);
			const inRun = (await chunkEnds(bytes, SHEET)).filter(
				(at) => (data[at] as number) >= 0x80,
			);
			expect(inRun.length).toBeGreaterThan(0);

			const expected = await buffered(bytes);
			expect(expected).toHaveLength(1_200);
			expect(expected[999]?.cells[0]?.value).toBe(`${BOM.repeat(40)}1000`);
			expect(await streamed(bytes)).toEqual(expected);
		},
	);

	it("streams a small stored sheet in one slice, as before", async () => {
		const rows = [
			[1, "a"],
			[2, "b"],
		];
		const bytes = await storedCopy(await writeXlsx({ sheets: [{ name: "Data", rows }] }));
		expect(openZip(bytes).entries.get(SHEET)?.compressedSize).toBeLessThan(16 * 1024);
		const expected = await buffered(bytes);
		expect(expected.map((row) => row.cells.map((cell) => cell.value))).toEqual(rows);
		expect(await streamed(bytes)).toEqual(expected);
	});
});
