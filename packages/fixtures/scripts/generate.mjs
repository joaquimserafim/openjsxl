// Generate the programmatic test fixtures for openjsxl from declarative specs.
//
// Each fixture is a spec passed to buildWorkbook (see ./build-workbook.mjs) — no hand-written
// OOXML. Output is deterministic, so re-running produces byte-identical files. Run: pnpm fixtures
//
// Real-world fixtures produced by Excel, LibreOffice, and Google Sheets live under ./data too
// and are NOT generated here — see data/README.md and ../THIRD_PARTY.md.

import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { buildWorkbook, packParts } from "./build-workbook.mjs";

// One sheet covering every value type the reader must handle, matching the original basic.xlsx:
// a shared string, a number, a date-styled serial (built-in numFmtId 14), a boolean, a cached
// formula, and a second row.
const fixtures = [
	{
		file: "basic.xlsx",
		spec: {
			sheets: [
				{
					name: "Sheet1",
					dimension: "A1:E2",
					cells: [
						{ ref: "A1", text: "hello" },
						{ ref: "B1", number: 42 },
						{ ref: "C1", serial: 43831, numFmtId: 14 },
						{ ref: "D1", bool: true },
						{ ref: "E1", formula: "B1*2", number: 84 },
						{ ref: "A2", text: "world" },
						{ ref: "B2", number: 3.14159 },
					],
				},
			],
		},
	},
	// A minimal workbook: numbers and a boolean only, so it carries NO sharedStrings.xml and NO
	// styles.xml — exercises the reader tolerating missing optional parts (F2.4b).
	{
		file: "minimal.xlsx",
		spec: {
			sheets: [
				{
					name: "Sheet1",
					cells: [
						{ ref: "A1", number: 1 },
						{ ref: "B1", bool: true },
					],
				},
			],
		},
	},
	// Style inheritance (#22): a format set at the column and row level, NOT on the cell.
	//  - Column B carries a date format; B1/B2 omit their own `s` → must read as dates.
	//  - Row 3 is a customFormat percent row; A3 omits `s` → inherits the percent.
	//  - B3 sets its own date `s` → the cell wins over row 3's percent (precedence check).
	//  - C1 is an explicit per-cell date (control); A1 is an unstyled plain number.
	{
		file: "col-row-styles.xlsx",
		spec: {
			sheets: [
				{
					name: "Sheet1",
					dimension: "A1:C3",
					columns: [{ min: 2, max: 2, numFmtId: 14 }],
					rowStyles: { 3: { numFmt: "0.00%" } },
					cells: [
						{ ref: "A1", number: 1 },
						{ ref: "B1", serial: 43831 },
						{ ref: "C1", serial: 43831, numFmtId: 14 },
						{ ref: "B2", serial: 44000 },
						{ ref: "A3", number: 0.5 },
						{ ref: "B3", serial: 44100, numFmtId: 14 },
					],
				},
			],
		},
	},
];

// Deliberately-broken packages for the error paths (F2.4b). Valid ZIPs, invalid OOXML.
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOC =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const XMLD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const broken = [
	// Package relationships name no officeDocument → XlsxError('not-xlsx').
	{
		file: "broken-no-officedoc.xlsx",
		parts: [
			{
				name: "_rels/.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"></Relationships>`,
			},
		],
	},
	// officeDocument points at xl/workbook.xml, but that part is absent → XlsxError('missing-part').
	{
		file: "broken-no-workbook.xlsx",
		parts: [
			{
				name: "_rels/.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOC}" Target="xl/workbook.xml"/></Relationships>`,
			},
		],
	},
	// Two central-directory entries with the same name → the duplicate is rejected (corrupt-zip).
	{
		file: "edge-duplicate-entry.xlsx",
		parts: [
			{ name: "x.xml", xml: "a" },
			{ name: "x.xml", xml: "b" },
		],
	},
	// A directory placeholder entry (name ends in '/') is skipped; the real part remains.
	{
		file: "edge-with-directory.xlsx",
		parts: [
			{ name: "sub/", xml: "" },
			{ name: "keep.xml", xml: "hi" },
		],
	},
	// A VALID package whose worksheet holds an adversarial cell ref: a column made of 300 'A's,
	// far past Excel's XFD limit. columnToIndex must reject it (rather than overflow to a
	// non-integer) so the reader falls back to positional addressing instead of throwing a bare
	// Error out of the read path. Should open cleanly and yield one row of two cells. (F2.4e —
	// regression for a bare-throw found by adversarial review, not by the seeded fuzz.)
	{
		file: "edge-overflow-col.xlsx",
		parts: [
			{
				name: "_rels/.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOC}" Target="xl/workbook.xml"/></Relationships>`,
			},
			{
				name: "xl/workbook.xml",
				xml: `${XMLD}\n<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
			},
			{
				name: "xl/_rels/workbook.xml.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
			},
			{
				name: "xl/worksheets/sheet1.xml",
				xml: `${XMLD}\n<worksheet><sheetData><row r="1"><c r="${"A".repeat(300)}1"><v>1</v></c><c><v>2</v></c></row></sheetData></worksheet>`,
			},
		],
	},
];

// Hand-crafted image fixtures (F6.2). openpyxl won't produce a drawing where two pictures share one
// media part, nor the degrade cases, so this package is authored part-by-part with packParts. Media
// bytes are OPAQUE to the reader (never decoded; mime comes from the extension), so short recognizable
// byte strings stand in for real images. Anchors use the DEFAULT (unprefixed) spreadsheetDrawing
// namespace exactly as Excel/openpyxl emit, with a:/r: only on the blip.
const SS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SHARED_MEDIA = new TextEncoder().encode("openjsxl-shared-media-bytes");

const pic = (id, name, embed) =>
	`<pic><nvPicPr><cNvPr id="${id}" name="${name}"/><cNvPicPr/></nvPicPr>` +
	`<blipFill><a:blip r:embed="${embed}"/><a:stretch><a:fillRect/></a:stretch></blipFill>` +
	`<spPr><a:prstGeom prst="rect"/></spPr></pic><clientData/>`;
const from = (col, row) =>
	`<from><col>${col}</col><colOff>0</colOff><row>${row}</row><rowOff>0</rowOff></from>`;
const to = (col, row) =>
	`<to><col>${col}</col><colOff>0</colOff><row>${row}</row><rowOff>0</rowOff></to>`;

const imagesEdgeDrawing =
	`${XMLD}\n<wsDr xmlns="${SS_DRAW}" xmlns:a="${A_MAIN}" xmlns:r="${R_NS}">` +
	// A: oneCellAnchor, blip rId1 → shared.png (kept).
	`<oneCellAnchor>${from(0, 0)}<ext cx="100" cy="100"/>${pic(1, "A", "rId1")}</oneCellAnchor>` +
	// B: twoCellAnchor, blip rId1 → shared.png (kept; shares A's media buffer).
	`<twoCellAnchor>${from(2, 2)}${to(4, 4)}${pic(2, "B", "rId1")}</twoCellAnchor>` +
	// C: absoluteAnchor with a picture → SKIPPED (geometry isn't cell-relative).
	`<absoluteAnchor><pos x="0" y="0"/><ext cx="100" cy="100"/>${pic(3, "C", "rId1")}</absoluteAnchor>` +
	// D: a SHAPE, not a picture → SKIPPED (no <pic>/<a:blip>).
	`<twoCellAnchor>${from(0, 6)}${to(1, 7)}<sp><nvSpPr><cNvPr id="4" name="D"/><cNvSpPr/></nvSpPr><spPr/></sp><clientData/></twoCellAnchor>` +
	// E: blip rId2 → ../media/gone.png, which is ABSENT from the zip → SKIPPED.
	`<oneCellAnchor>${from(0, 9)}<ext cx="50" cy="50"/>${pic(5, "E", "rId2")}</oneCellAnchor>` +
	// F: blip rId3, which is NOT in the drawing rels → SKIPPED.
	`<oneCellAnchor>${from(0, 11)}<ext cx="50" cy="50"/>${pic(6, "F", "rId3")}</oneCellAnchor>` +
	"</wsDr>";

const crafted = [
	{
		file: "images-edge.xlsx",
		parts: [
			{
				name: "[Content_Types].xml",
				xml: `${XMLD}\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
			},
			{
				name: "_rels/.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOC}" Target="xl/workbook.xml"/></Relationships>`,
			},
			{
				name: "xl/workbook.xml",
				xml: `${XMLD}\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R_NS}"><sheets><sheet name="Pics" sheetId="1" r:id="rId1"/></sheets></workbook>`,
			},
			{
				name: "xl/_rels/workbook.xml.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${R_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
			},
			{
				name: "xl/worksheets/sheet1.xml",
				xml: `${XMLD}\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>pics</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`,
			},
			{
				name: "xl/worksheets/_rels/sheet1.xml.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
			},
			{ name: "xl/drawings/drawing1.xml", xml: imagesEdgeDrawing },
			{
				// rId1 → shared.png (present); rId2 → gone.png (NOT packed below). rId3 is unreferenced.
				name: "xl/drawings/_rels/drawing1.xml.rels",
				xml: `${XMLD}\n<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${R_NS}/image" Target="../media/shared.png"/><Relationship Id="rId2" Type="${R_NS}/image" Target="../media/gone.png"/></Relationships>`,
			},
			{ name: "xl/media/shared.png", data: SHARED_MEDIA },
		],
	},
];

// ── Hand-crafted OpenDocument spreadsheet (.ods) fixtures (F7.1) ────────────────────────────────
// An .ods is a ZIP whose FIRST entry is a STORED `mimetype`; sheets all live in one content.xml.
// packParts stores every entry in array order, so putting `mimetype` first yields an ODF-conformant
// container. These cover the shapes a real producer (LibreOffice) would emit that odfpy won't easily
// reproduce — the repeat-to-grid-edge tail (the "repeat bomb"), covered-cell merges, a hidden and an
// empty sheet — plus the typed reject cases (encrypted, wrong document type, no content.xml).
const ODF_NS =
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
	'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
	'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
	'xmlns:xlink="http://www.w3.org/1999/xlink"';
const MIMETYPE_ODS = "application/vnd.oasis.opendocument.spreadsheet";
const odfManifest = (mediaType) =>
	`${XMLD}\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
	`<manifest:file-entry manifest:full-path="/" manifest:media-type="${mediaType}"/>` +
	`<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
const odfContent = (body) =>
	`${XMLD}\n<office:document-content ${ODF_NS} office:version="1.3"><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`;

// The edge sheet: row 1 has 6 real columns (one empty, one a hyperlink) then an empty repeat to the
// 16 384-column edge; row 2 fills a value across 3 columns; rows 3-4 are a 2×2 covered-cell merge at
// A3; row 5 repeats an empty row to the 2^20 grid edge (the bomb — must cost O(1) to read).
const odsEdgeBody =
	'<table:table table:name="Data">' +
	"<table:table-row>" +
	'<table:table-cell office:value-type="string"><text:p>hello</text:p></table:table-cell>' +
	'<table:table-cell office:value-type="float" office:value="42"><text:p>42</text:p></table:table-cell>' +
	'<table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>' +
	'<table:table-cell office:value-type="date" office:date-value="2024-01-15"><text:p>2024-01-15</text:p></table:table-cell>' +
	"<table:table-cell/>" +
	'<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com/">link</text:a></text:p></table:table-cell>' +
	'<table:table-cell table:number-columns-repeated="16378"/>' +
	"</table:table-row>" +
	"<table:table-row>" +
	'<table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="3"><text:p>7</text:p></table:table-cell>' +
	'<table:table-cell table:number-columns-repeated="16381"/>' +
	"</table:table-row>" +
	"<table:table-row>" +
	'<table:table-cell office:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>merged</text:p></table:table-cell>' +
	"<table:covered-table-cell/>" +
	'<table:table-cell table:number-columns-repeated="16382"/>' +
	"</table:table-row>" +
	"<table:table-row>" +
	'<table:covered-table-cell table:number-columns-repeated="2"/>' +
	'<table:table-cell table:number-columns-repeated="16382"/>' +
	"</table:table-row>" +
	'<table:table-row table:number-rows-repeated="1048571">' +
	'<table:table-cell table:number-columns-repeated="16384"/>' +
	"</table:table-row>" +
	"</table:table>" +
	'<table:table table:name="Hidden" table:visibility="collapse">' +
	'<table:table-row><table:table-cell office:value-type="string"><text:p>secret</text:p></table:table-cell></table:table-row>' +
	"</table:table>" +
	'<table:table table:name="Blank"/>';

const ods = [
	{
		file: "ods-edge.ods",
		parts: [
			{ name: "mimetype", data: new TextEncoder().encode(MIMETYPE_ODS) },
			{ name: "content.xml", xml: odfContent(odsEdgeBody) },
			{ name: "META-INF/manifest.xml", xml: odfManifest(MIMETYPE_ODS) },
		],
	},
	{
		// A valid ZIP whose manifest declares encryption — openOds must refuse it typed (unsupported).
		file: "ods-encrypted.ods",
		parts: [
			{ name: "mimetype", data: new TextEncoder().encode(MIMETYPE_ODS) },
			{ name: "content.xml", xml: odfContent('<table:table table:name="S"/>') },
			{
				name: "META-INF/manifest.xml",
				xml:
					`${XMLD}\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
					'<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">' +
					'<manifest:encryption-data manifest:checksum-type="SHA1"/></manifest:file-entry></manifest:manifest>',
			},
		],
	},
	{
		// A valid ODF ZIP that is a TEXT document, not a spreadsheet — openOds must reject the mimetype.
		file: "ods-not-spreadsheet.ods",
		parts: [
			{
				name: "mimetype",
				data: new TextEncoder().encode("application/vnd.oasis.opendocument.text"),
			},
			{ name: "content.xml", xml: `${XMLD}\n<office:document-content/>` },
		],
	},
	{
		// A ZIP with the right mimetype but NO content.xml — openOds must fail with missing-part.
		file: "ods-no-content.ods",
		parts: [{ name: "mimetype", data: new TextEncoder().encode(MIMETYPE_ODS) }],
	},
];

const dataDir = new URL("../data/", import.meta.url);
await mkdir(dataDir, { recursive: true });

for (const { file, spec } of fixtures) {
	const outUrl = new URL(file, dataDir);
	const archive = buildWorkbook(spec);
	await writeFile(outUrl, archive);
	console.log(`wrote ${fileURLToPath(outUrl)} (${archive.length} bytes)`);
}

for (const { file, parts } of broken) {
	const outUrl = new URL(file, dataDir);
	await writeFile(outUrl, packParts(parts));
	console.log(`wrote ${fileURLToPath(outUrl)} (broken fixture)`);
}

for (const { file, parts } of crafted) {
	const outUrl = new URL(file, dataDir);
	await writeFile(outUrl, packParts(parts));
	console.log(`wrote ${fileURLToPath(outUrl)} (crafted fixture)`);
}

for (const { file, parts } of ods) {
	const outUrl = new URL(file, dataDir);
	await writeFile(outUrl, packParts(parts));
	console.log(`wrote ${fileURLToPath(outUrl)} (crafted .ods fixture)`);
}
