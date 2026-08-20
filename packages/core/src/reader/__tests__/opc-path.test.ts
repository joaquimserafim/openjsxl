import { describe, expect, it } from "vitest";
import { directoryOf, relsPathFor } from "../opc-path";

describe("OPC part paths", () => {
	it("returns a part's containing directory", () => {
		expect(directoryOf("xl/workbook.xml")).toBe("xl");
		expect(directoryOf("xl/worksheets/sheet1.xml")).toBe("xl/worksheets");
		expect(directoryOf("workbook.xml")).toBe("");
	});

	it("builds the adjacent relationships-part path", () => {
		expect(relsPathFor("xl/workbook.xml")).toBe("xl/_rels/workbook.xml.rels");
		expect(relsPathFor("xl/worksheets/sheet1.xml")).toBe("xl/worksheets/_rels/sheet1.xml.rels");
		expect(relsPathFor("workbook.xml")).toBe("_rels/workbook.xml.rels");
	});
});
