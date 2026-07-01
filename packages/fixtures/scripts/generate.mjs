// Generate the programmatic test fixtures for openjsxl from declarative specs.
//
// Each fixture is a spec passed to buildWorkbook (see ./build-workbook.mjs) — no hand-written
// OOXML. Output is deterministic, so re-running produces byte-identical files. Run: pnpm fixtures
//
// Real-world fixtures produced by Excel, LibreOffice, and Google Sheets live under ./data too
// and are NOT generated here — see data/README.md and ../THIRD_PARTY.md.

import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { buildWorkbook, packParts } from './build-workbook.mjs'

// One sheet covering every value type the reader must handle, matching the original basic.xlsx:
// a shared string, a number, a date-styled serial (built-in numFmtId 14), a boolean, a cached
// formula, and a second row.
const fixtures = [
	{
		file: 'basic.xlsx',
		spec: {
			sheets: [
				{
					name: 'Sheet1',
					dimension: 'A1:E2',
					cells: [
						{ ref: 'A1', text: 'hello' },
						{ ref: 'B1', number: 42 },
						{ ref: 'C1', serial: 43831, numFmtId: 14 },
						{ ref: 'D1', bool: true },
						{ ref: 'E1', formula: 'B1*2', number: 84 },
						{ ref: 'A2', text: 'world' },
						{ ref: 'B2', number: 3.14159 },
					],
				},
			],
		},
	},
	// A minimal workbook: numbers and a boolean only, so it carries NO sharedStrings.xml and NO
	// styles.xml — exercises the reader tolerating missing optional parts (F2.4b).
	{
		file: 'minimal.xlsx',
		spec: {
			sheets: [
				{
					name: 'Sheet1',
					cells: [
						{ ref: 'A1', number: 1 },
						{ ref: 'B1', bool: true },
					],
				},
			],
		},
	},
]

// Deliberately-broken packages for the error paths (F2.4b). Valid ZIPs, invalid OOXML.
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const XMLD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const broken = [
	// Package relationships name no officeDocument → XlsxError('not-xlsx').
	{
		file: 'broken-no-officedoc.xlsx',
		parts: [
			{
				name: '_rels/.rels',
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"></Relationships>`,
			},
		],
	},
	// officeDocument points at xl/workbook.xml, but that part is absent → XlsxError('missing-part').
	{
		file: 'broken-no-workbook.xlsx',
		parts: [
			{
				name: '_rels/.rels',
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOC}" Target="xl/workbook.xml"/></Relationships>`,
			},
		],
	},
	// Two central-directory entries with the same name → the duplicate is rejected (corrupt-zip).
	{
		file: 'edge-duplicate-entry.xlsx',
		parts: [
			{ name: 'x.xml', xml: 'a' },
			{ name: 'x.xml', xml: 'b' },
		],
	},
	// A directory placeholder entry (name ends in '/') is skipped; the real part remains.
	{
		file: 'edge-with-directory.xlsx',
		parts: [
			{ name: 'sub/', xml: '' },
			{ name: 'keep.xml', xml: 'hi' },
		],
	},
	// A VALID package whose worksheet holds an adversarial cell ref: a column made of 300 'A's,
	// far past Excel's XFD limit. columnToIndex must reject it (rather than overflow to a
	// non-integer) so the reader falls back to positional addressing instead of throwing a bare
	// Error out of the read path. Should open cleanly and yield one row of two cells. (F2.4e —
	// regression for a bare-throw found by adversarial review, not by the seeded fuzz.)
	{
		file: 'edge-overflow-col.xlsx',
		parts: [
			{
				name: '_rels/.rels',
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOC}" Target="xl/workbook.xml"/></Relationships>`,
			},
			{
				name: 'xl/workbook.xml',
				xml: `${XMLD}\n<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
			},
			{
				name: 'xl/_rels/workbook.xml.rels',
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
			},
			{
				name: 'xl/worksheets/sheet1.xml',
				xml: `${XMLD}\n<worksheet><sheetData><row r="1"><c r="${'A'.repeat(300)}1"><v>1</v></c><c><v>2</v></c></row></sheetData></worksheet>`,
			},
		],
	},
]

const dataDir = new URL('../data/', import.meta.url)
await mkdir(dataDir, { recursive: true })

for (const { file, spec } of fixtures) {
	const outUrl = new URL(file, dataDir)
	const archive = buildWorkbook(spec)
	await writeFile(outUrl, archive)
	console.log(`wrote ${fileURLToPath(outUrl)} (${archive.length} bytes)`)
}

for (const { file, parts } of broken) {
	const outUrl = new URL(file, dataDir)
	await writeFile(outUrl, packParts(parts))
	console.log(`wrote ${fileURLToPath(outUrl)} (broken fixture)`)
}
