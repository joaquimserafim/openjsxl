import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import { openZip } from "../../zip";
import { resolveHyperlinks } from "../xlsb";

// The xlsb hyperlink resolver: each sheet hyperlink's rel id through the sheet's own rels part,
// which is claimed for the open first (F11.95).

const SHEET = "xl/worksheets/sheet1.bin";
const RELS = "xl/worksheets/_rels/sheet1.bin.rels";
const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const archive = openZip(
	await writeZip([
		{ name: SHEET, data: new Uint8Array(0) },
		{
			name: RELS,
			data: new TextEncoder().encode(
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					`<Relationship Id="rId1" Type="${HYPERLINK}" Target="https://example.com/" TargetMode="External"/>` +
					`<Relationship Id="rId2" Type="${HYPERLINK}" Target="sheet2.bin"/>` +
					"</Relationships>",
			),
		},
	]),
);

describe("resolveHyperlinks", () => {
	it("resolves nothing for a sheet without hyperlinks", async () => {
		expect(await resolveHyperlinks(archive, SHEET, [], new Set())).toEqual([]);
	});

	it("resolves external targets, in-workbook locations and unknown rel ids", async () => {
		const claimed = new Set<string>();
		expect(
			await resolveHyperlinks(
				archive,
				SHEET,
				[
					{ ref: "A1", relId: "rId1" },
					{ ref: "B1", relId: "rId2" },
					{ ref: "C1", relId: "rId9" },
				],
				claimed,
			),
		).toEqual([
			{ ref: "A1", target: "https://example.com/" },
			{ ref: "B1", location: "xl/worksheets/sheet2.bin" },
			{ ref: "C1" },
		]);
		expect([...claimed]).toEqual([RELS]);
	});

	it("keeps only the refs when this open already read the rels part", async () => {
		expect(
			await resolveHyperlinks(
				archive,
				SHEET,
				[{ ref: "A1", relId: "rId1" }],
				new Set([RELS]),
			),
		).toEqual([{ ref: "A1" }]);
	});
});
