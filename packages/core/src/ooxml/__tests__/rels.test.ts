import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { openZip } from "../../zip";
import { parseRels, resolveTarget } from "../rels";

describe("parseRels", () => {
	it("parses relationships into an id-keyed map", () => {
		const xml = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://x/officeDocument" Target="xl/workbook.xml"/>
	<Relationship Id="rId2" Type="http://x/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`;
		const rels = parseRels(xml);
		expect(rels.size).toBe(2);
		expect(rels.get("rId1")).toEqual({
			id: "rId1",
			type: "http://x/officeDocument",
			target: "xl/workbook.xml",
			targetMode: "Internal",
		});
		expect(rels.get("rId2")?.targetMode).toBe("External");
	});

	it("ignores non-Relationship elements and entries missing Id or Target", () => {
		const xml =
			'<Relationships><Relationship Id="rId1"/><Other Id="x" Target="y"/></Relationships>';
		expect(parseRels(xml).size).toBe(0);
	});
});

describe("resolveTarget", () => {
	it("resolves targets against the part base directory", () => {
		expect(resolveTarget("", "xl/workbook.xml")).toBe("xl/workbook.xml");
		expect(resolveTarget("xl", "worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml");
		expect(resolveTarget("xl/worksheets", "../media/image1.png")).toBe("xl/media/image1.png");
		expect(resolveTarget("xl", "/xl/styles.xml")).toBe("xl/styles.xml");
	});
});

describe("rels — real basic.xlsx", () => {
	it("walks the package rels to the workbook and on to its sheet", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"));
		const dec = new TextDecoder();

		const pkgRels = parseRels(dec.decode(await zip.read("_rels/.rels")));
		const office = [...pkgRels.values()].find((r) => r.type.endsWith("/officeDocument"));
		expect(resolveTarget("", office?.target ?? "")).toBe("xl/workbook.xml");

		const wbRels = parseRels(dec.decode(await zip.read("xl/_rels/workbook.xml.rels")));
		const sheet = [...wbRels.values()].find((r) => r.type.endsWith("/worksheet"));
		expect(resolveTarget("xl", sheet?.target ?? "")).toBe("xl/worksheets/sheet1.xml");
	});
});
