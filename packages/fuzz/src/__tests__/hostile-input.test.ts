import { openXlsx, streamSheetRows, XlsxError } from "@openjsxl/core";
import { describe, expect, it } from "vitest";

// F9.7 — hostile-input hardening, driven end-to-end through the PUBLIC readers under the DEFAULT
// guards (no options passed). The property: a decompression bomb or a lying-size part makes the
// reader fail TYPED (`part-too-large`) quickly, never OOM or hang. Complements the layer-level
// units (zip/__tests__/central-directory + xml/__tests__/stream); here the whole open→read path
// is exercised on a real OPC package with a hostile worksheet.

const enc = new TextEncoder();
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// ── a minimal multi-entry zip packer (mixed stored / deflate), enough to re-pack an xlsx ────────
const CRC_TABLE = (() => {
	const t = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		t[n] = c >>> 0;
	}
	return t;
})();
function crc32(bytes: Uint8Array): number {
	let crc = 0xffffffff;
	for (let i = 0; i < bytes.length; i++) {
		crc = ((CRC_TABLE[(crc ^ (bytes[i] as number)) & 0xff] as number) ^ (crc >>> 8)) >>> 0;
	}
	return (crc ^ 0xffffffff) >>> 0;
}
const u16 = (n: number): number[] => [n & 0xff, (n >>> 8) & 0xff];
const u32 = (n: number): number[] => [
	n & 0xff,
	(n >>> 8) & 0xff,
	(n >>> 16) & 0xff,
	(n >>> 24) & 0xff,
];

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	const s = new Blob([data as unknown as BlobPart])
		.stream()
		.pipeThrough(new CompressionStream("deflate-raw"));
	return new Uint8Array(await new Response(s).arrayBuffer());
}

interface PackEntry {
	name: string;
	bytes: Uint8Array;
	deflate: boolean;
	/** override the central-directory uncompressedSize (a lie, to probe the size guard) */
	fakeUncompressed?: number;
}

async function packZip(entries: PackEntry[]): Promise<Uint8Array> {
	const chunks: number[] = [];
	const central: number[] = [];
	let offset = 0;
	for (const e of entries) {
		const nameBytes = enc.encode(e.name);
		const raw = e.bytes;
		const payload = e.deflate ? await deflateRaw(raw) : raw;
		const crc = crc32(raw);
		const uncompressed = e.fakeUncompressed ?? raw.length;
		const method = e.deflate ? 8 : 0;
		const local = [
			...u32(0x04034b50),
			...u16(20),
			...u16(0),
			...u16(method),
			...u16(0),
			...u16(0),
			...u32(crc),
			...u32(payload.length),
			...u32(uncompressed),
			...u16(nameBytes.length),
			...u16(0),
			...nameBytes,
		];
		chunks.push(...local, ...payload);
		central.push(
			...u32(0x02014b50),
			...u16(20),
			...u16(20),
			...u16(0),
			...u16(method),
			...u16(0),
			...u16(0),
			...u32(crc),
			...u32(payload.length),
			...u32(uncompressed),
			...u16(nameBytes.length),
			...u16(0),
			...u16(0),
			...u16(0),
			...u16(0),
			...u32(0),
			...u32(offset),
			...nameBytes,
		);
		offset += local.length + payload.length;
	}
	const eocd = [
		...u32(0x06054b50),
		...u16(0),
		...u16(0),
		...u16(entries.length),
		...u16(entries.length),
		...u32(central.length),
		...u32(offset),
		...u16(0),
	];
	return Uint8Array.from([...chunks, ...central, ...eocd]);
}

// Pack a minimal valid OPC package for a one-sheet "S" workbook, replacing its worksheet with a
// hostile one. Scaffolding parts are stored verbatim so openXlsx reaches the worksheet.
async function repackWithHostileSheet(
	hostileSheetXml: string,
	opts?: { deflate?: boolean; fakeUncompressed?: number },
): Promise<Uint8Array> {
	const entries: PackEntry[] = baseParts().map((p) =>
		p.name === "xl/worksheets/sheet1.xml"
			? {
					name: p.name,
					bytes: enc.encode(hostileSheetXml),
					deflate: opts?.deflate ?? true,
					...(opts?.fakeUncompressed !== undefined
						? { fakeUncompressed: opts.fakeUncompressed }
						: {}),
				}
			: { name: p.name, bytes: enc.encode(p.xml), deflate: false },
	);
	return packZip(entries);
}

// The fixed OPC scaffolding for a one-sheet "S" workbook — enough for openXlsx to reach the
// worksheet. Kept inline so the test needs no raw-zip reader.
function baseParts(): { name: string; xml: string }[] {
	return [
		{
			name: "[Content_Types].xml",
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
		},
		{
			name: "_rels/.rels",
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		},
		{
			name: "xl/workbook.xml",
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		},
		{
			name: "xl/_rels/workbook.xml.rels",
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		},
		{ name: "xl/worksheets/sheet1.xml", xml: "" }, // replaced by the hostile sheet
	];
}

// A worksheet part that is a valid element tree but a decompression BOMB: a ~16 MiB run of one
// character (deflates to a few KB → ~thousands:1 ratio).
function bombSheet(): string {
	const filler = "A".repeat(16 * 1024 * 1024);
	return `<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${filler}</t></is></c></row></sheetData></worksheet>`;
}

describe("F9.7 — hostile input fails typed (not OOM/hang) under the DEFAULT guards", () => {
	it("openXlsx rejects a decompression-bomb worksheet typed, quickly", async () => {
		const bytes = await repackWithHostileSheet(bombSheet());
		const t0 = performance.now();
		const err = await openXlsx(bytes)
			.then(() => undefined)
			.catch((e) => e);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("part-too-large");
		expect(performance.now() - t0).toBeLessThan(5000); // aborts early — no full 16 MiB expand-then-hang
	});

	it("streamSheetRows rejects a bomb worksheet typed, without accumulating it", async () => {
		const bytes = await repackWithHostileSheet(bombSheet());
		const err = await (async () => {
			try {
				for await (const _row of streamSheetRows(bytes)) {
					// The stream must abort at the ratio/floor cap before yielding a full sheet.
				}
				return undefined;
			} catch (e) {
				return e;
			}
		})();
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("part-too-large");
	});

	it("rejects a part whose central uncompressedSize lies enormous, before inflating", async () => {
		// A modest real sheet, but the central directory declares a ~3 GiB output — over the 2 GiB
		// absolute default. The guard must refuse at locate(), never attempt the allocation.
		const sheet = `<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;
		const bytes = await repackWithHostileSheet(sheet, {
			deflate: true,
			fakeUncompressed: 3 * 1024 * 1024 * 1024,
		});
		const err = await openXlsx(bytes)
			.then(() => undefined)
			.catch((e) => e);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("part-too-large");
	});

	it("the same package with a NORMAL worksheet still opens (guard doesn't false-positive)", async () => {
		const sheet = `<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>`;
		const bytes = await repackWithHostileSheet(sheet);
		const wb = await openXlsx(bytes);
		expect(wb.sheet("S").cell("A1").value).toBe(42);
	});
});
