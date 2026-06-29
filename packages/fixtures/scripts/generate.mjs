// Generate the programmatic test fixtures for openjsxl.
//
// Builds a real, valid .xlsx by writing the minimal OOXML parts and packing them into
// a ZIP using STORED (uncompressed) entries — so this script needs no zip or deflate
// dependency, only a small CRC32. Output is deterministic (fixed DOS timestamps), so
// re-running produces byte-identical files. Run with: pnpm fixtures
//
// Real-world fixtures produced by Excel, LibreOffice, and Google Sheets should be
// committed directly under ./data alongside whatever this generates (see data/README.md).

import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

const encoder = new TextEncoder()

const CRC_TABLE = (() => {
	const table = new Uint32Array(256)
	for (let n = 0; n < 256; n++) {
		let c = n
		for (let k = 0; k < 8; k++) {
			c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		}
		table[n] = c >>> 0
	}
	return table
})()

function crc32(bytes) {
	let crc = 0xffffffff
	for (let i = 0; i < bytes.length; i++) {
		crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8)
	}
	return (crc ^ 0xffffffff) >>> 0
}

const u16 = (n) => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff])
const u32 = (n) =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff])

function concat(parts) {
	let total = 0
	for (const part of parts) total += part.length
	const out = new Uint8Array(total)
	let offset = 0
	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}
	return out
}

// Fixed DOS date (1980-01-01) and time (00:00) for deterministic output.
const DOS_TIME = 0
const DOS_DATE = 0x0021

function zipStore(files) {
	const local = []
	const central = []
	let offset = 0

	for (const file of files) {
		const name = encoder.encode(file.name)
		const { data } = file
		const crc = crc32(data)
		const size = data.length

		const header = concat([
			u32(0x04034b50), // local file header signature
			u16(20), // version needed
			u16(0), // flags
			u16(0), // method: 0 = stored
			u16(DOS_TIME),
			u16(DOS_DATE),
			u32(crc),
			u32(size), // compressed size
			u32(size), // uncompressed size
			u16(name.length),
			u16(0), // extra length
			name,
		])
		local.push(header, data)

		central.push(
			concat([
				u32(0x02014b50), // central directory header signature
				u16(20), // version made by
				u16(20), // version needed
				u16(0), // flags
				u16(0), // method
				u16(DOS_TIME),
				u16(DOS_DATE),
				u32(crc),
				u32(size),
				u32(size),
				u16(name.length),
				u16(0), // extra length
				u16(0), // comment length
				u16(0), // disk number start
				u16(0), // internal attributes
				u32(0), // external attributes
				u32(offset), // local header offset
				name,
			]),
		)

		offset += header.length + data.length
	}

	const directory = concat(central)
	const eocd = concat([
		u32(0x06054b50), // end of central directory signature
		u16(0), // this disk
		u16(0), // disk with central directory
		u16(files.length), // entries on this disk
		u16(files.length), // total entries
		u32(directory.length),
		u32(offset), // central directory offset
		u16(0), // comment length
	])

	return concat([...local, directory, eocd])
}

// Minimal workbook: one sheet with a string, number, date, boolean, a cached formula,
// and a second row. [Content_Types].xml is intentionally the first entry.
const parts = {
	'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
\t<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
\t<Default Extension="xml" ContentType="application/xml"/>
\t<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
\t<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
\t<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
\t<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
	'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
\t<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
	'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
\t<sheets>
\t\t<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
\t</sheets>
</workbook>`,
	'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
\t<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
\t<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
\t<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
	'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
\t<dimension ref="A1:E2"/>
\t<sheetData>
\t\t<row r="1">
\t\t\t<c r="A1" t="s"><v>0</v></c>
\t\t\t<c r="B1"><v>42</v></c>
\t\t\t<c r="C1" s="1"><v>43831</v></c>
\t\t\t<c r="D1" t="b"><v>1</v></c>
\t\t\t<c r="E1"><f>B1*2</f><v>84</v></c>
\t\t</row>
\t\t<row r="2">
\t\t\t<c r="A2" t="s"><v>1</v></c>
\t\t\t<c r="B2"><v>3.14159</v></c>
\t\t</row>
\t</sheetData>
</worksheet>`,
	'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
\t<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
\t<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
\t<borders count="1"><border/></borders>
\t<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
\t<cellXfs count="2">
\t\t<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
\t\t<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
\t</cellXfs>
</styleSheet>`,
	'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
\t<si><t>hello</t></si>
\t<si><t>world</t></si>
</sst>`,
}

const files = Object.entries(parts).map(([name, xml]) => ({
	name,
	data: encoder.encode(xml),
}))

const dataDir = new URL('../data/', import.meta.url)
await mkdir(dataDir, { recursive: true })

const outUrl = new URL('basic.xlsx', dataDir)
const archive = zipStore(files)
await writeFile(outUrl, archive)

console.log(`wrote ${fileURLToPath(outUrl)} (${archive.length} bytes, ${files.length} parts)`)
