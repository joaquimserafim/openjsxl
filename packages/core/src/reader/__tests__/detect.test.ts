import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import { detectSpreadsheetFormat } from "../detect";

// F7.4 — `detectSpreadsheetFormat` sniffs the container so a caller can route to the right opener.
// Real fixtures pin the four positive cases; crafted zips pin the OOXML variants (xlsm/xltx) and the
// negative cases (non-spreadsheet zip, corrupt/truncated, binary junk). CSV has no magic bytes, so
// its detection is a documented best-effort heuristic — any decodable text reads as `'csv'`.

const enc = new TextEncoder();

// Minimal `[Content_Types].xml` declaring one workbook part of the given content type.
function contentTypes(mainType: string): string {
	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		`<Override PartName="/xl/workbook.xml" ContentType="${mainType}"/></Types>`
	);
}

// Pack a one-part zip whose `[Content_Types].xml` names `mainType` — enough for the sniff to classify.
function ooxmlZip(mainType: string): Promise<Uint8Array> {
	return writeZip([{ name: "[Content_Types].xml", data: enc.encode(contentTypes(mainType)) }]);
}

describe("detectSpreadsheetFormat — real fixtures", () => {
	it("detects .xlsx", async () => {
		expect(await detectSpreadsheetFormat(await loadFixture("basic.xlsx"))).toBe("xlsx");
	});
	it("detects .xlsb", async () => {
		expect(await detectSpreadsheetFormat(await loadFixture("xlsb-basic.xlsb"))).toBe("xlsb");
	});
	it("detects .ods", async () => {
		expect(await detectSpreadsheetFormat(await loadFixture("odf-basic.ods"))).toBe("ods");
	});
	it("detects .csv", async () => {
		expect(await detectSpreadsheetFormat(await loadFixture("basic.csv"))).toBe("csv");
	});
});

describe("detectSpreadsheetFormat — OOXML content-type variants all read as xlsx", () => {
	const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
	const XLTX = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
	const XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
	const XLTM = "application/vnd.ms-excel.template.macroEnabled.main+xml";
	for (const [label, type] of [
		["xlsx", XLSX],
		["xltx", XLTX],
		["xlsm", XLSM],
		["xltm", XLTM],
	] as const) {
		it(`classifies a ${label} workbook part as xlsx`, async () => {
			expect(await detectSpreadsheetFormat(await ooxmlZip(type))).toBe("xlsx");
		});
	}
});

describe("detectSpreadsheetFormat — negatives", () => {
	it("returns undefined for an empty input", async () => {
		expect(await detectSpreadsheetFormat(new Uint8Array(0))).toBeUndefined();
	});

	it("returns undefined for binary junk (NUL / invalid UTF-8)", async () => {
		expect(
			await detectSpreadsheetFormat(new Uint8Array([0x00, 0x01, 0x02, 0xff, 0xfe, 0x00])),
		).toBeUndefined();
	});

	it("returns undefined for a truncated zip (PK header, no central directory)", async () => {
		const full = await loadFixture("basic.xlsx");
		const truncated = full.subarray(0, Math.max(4, full.byteLength - 128));
		expect(await detectSpreadsheetFormat(truncated)).toBeUndefined();
	});

	it("returns undefined for a zip that is not a spreadsheet package (e.g. a .docx)", async () => {
		const docx = await ooxmlZip(
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
		);
		expect(await detectSpreadsheetFormat(docx)).toBeUndefined();
	});

	it("returns undefined for a plain zip with no OOXML/ODF markers", async () => {
		const zip = await writeZip([{ name: "notes.txt", data: enc.encode("just a zip") }]);
		expect(await detectSpreadsheetFormat(zip)).toBeUndefined();
	});
});

describe("detectSpreadsheetFormat — csv is a best-effort text heuristic", () => {
	it("classifies arbitrary decodable text as csv (documented breadth)", async () => {
		// No magic bytes exist for CSV, so any text — even prose with no delimiters — reads as csv.
		expect(await detectSpreadsheetFormat(enc.encode("hello world\nsecond line"))).toBe("csv");
	});

	it("tolerates a UTF-8 BOM", async () => {
		expect(await detectSpreadsheetFormat(enc.encode("﻿a,b,c\n1,2,3"))).toBe("csv");
	});

	it("accepts an ArrayBuffer", async () => {
		const buf = enc.encode("a,b\n1,2").buffer;
		expect(await detectSpreadsheetFormat(buf)).toBe("csv");
	});

	it("does not mistake text that starts with 'PK' for a zip", async () => {
		expect(await detectSpreadsheetFormat(enc.encode("PKG,version\ncore,1"))).toBe("csv");
	});
});

// An ODF content.xml with the given body; `office:spreadsheet` in the body marks a spreadsheet.
const odfContent = (body: string): string =>
	'<?xml version="1.0" encoding="UTF-8"?>' +
	'<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
	'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">' +
	`<office:body>${body}</office:body></office:document-content>`;
const SPREADSHEET_BODY = '<office:spreadsheet><table:table table:name="S"/></office:spreadsheet>';

describe("detectSpreadsheetFormat — mimetype-less ODS (openOds tolerates it; F7.4 review)", () => {
	it("classifies a mimetype-less ODS as ods via its content.xml spreadsheet body", async () => {
		// A few producers omit the optional `mimetype` entry; openOds reads such a file, so detect must
		// too (its verdict must never be stricter than its paired opener).
		const noMime = await writeZip([
			{ name: "content.xml", data: enc.encode(odfContent(SPREADSHEET_BODY)) },
		]);
		expect(await detectSpreadsheetFormat(noMime)).toBe("ods");
	});

	it("does NOT classify a mimetype-less TEXT document (.odt) as ods", async () => {
		const odt = await writeZip([
			{ name: "content.xml", data: enc.encode(odfContent("<office:text/>")) },
		]);
		expect(await detectSpreadsheetFormat(odt)).toBeUndefined();
	});
});

describe("detectSpreadsheetFormat — sniff reads a bounded prefix (bomb guard; F7.4 review)", () => {
	it("does not find a spreadsheet marker pushed past the 1 MiB sniff cap", async () => {
		// Detection streams each zip part and stops at a 1 MiB prefix, so a decompression bomb can't
		// force an unbounded inflate. The observable consequence, pinned here: a marker artificially
		// buried >1 MiB into content.xml is NOT seen (real ODS markers sit at the very top). The old
		// code read the whole part and would have returned "ods".
		const buried = odfContent(" ".repeat(1_200_000) + SPREADSHEET_BODY);
		const zip = await writeZip([{ name: "content.xml", data: enc.encode(buried) }]);
		expect(await detectSpreadsheetFormat(zip)).toBeUndefined();
	});
});
