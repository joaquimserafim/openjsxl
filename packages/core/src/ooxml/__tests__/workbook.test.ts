import { describe, expect, it } from "vitest";
import { parseWorkbook } from "../workbook";

describe("parseWorkbook", () => {
	it("lists sheets in order with name, r:id, and visibility", () => {
		const xml = `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<sheets>
		<sheet name="First" sheetId="1" r:id="rId1"/>
		<sheet name="Second" sheetId="2" r:id="rId2"/>
	</sheets>
</workbook>`;
		expect(parseWorkbook(xml).sheets).toEqual([
			{ name: "First", rid: "rId1", visible: true, state: "visible" },
			{ name: "Second", rid: "rId2", visible: true, state: "visible" },
		]);
	});

	it("marks hidden and very-hidden sheets not visible, with their state (F4.6)", () => {
		const xml =
			"<workbook><sheets>" +
			'<sheet name="A" r:id="rId1" state="visible"/>' +
			'<sheet name="B" r:id="rId2" state="hidden"/>' +
			'<sheet name="C" r:id="rId3" state="veryHidden"/>' +
			"</sheets></workbook>";
		expect(parseWorkbook(xml).sheets.map((s) => s.visible)).toEqual([true, false, false]);
		expect(parseWorkbook(xml).sheets.map((s) => s.state)).toEqual([
			"visible",
			"hidden",
			"veryHidden",
		]);
	});

	it("reads an unrecognized state as visible (the spec default)", () => {
		const xml =
			'<workbook><sheets><sheet name="A" r:id="rId1" state="banana"/></sheets></workbook>';
		expect(parseWorkbook(xml).sheets[0]?.state).toBe("visible");
		expect(parseWorkbook(xml).sheets[0]?.visible).toBe(true);
	});

	it("finds the relationship id under any namespace prefix", () => {
		const xml = '<workbook><sheets><sheet name="A" x:id="rId7"/></sheets></workbook>';
		expect(parseWorkbook(xml).sheets[0]?.rid).toBe("rId7");
	});

	it("skips sheets missing a name or relationship id", () => {
		const xml =
			'<workbook><sheets><sheet name="A"/><sheet r:id="rId2"/><sheet name="C" r:id="rId3"/></sheets></workbook>';
		expect(parseWorkbook(xml).sheets.map((s) => s.name)).toEqual(["C"]);
	});

	it("reads the date1904 flag from <workbookPr>, defaulting to false", () => {
		expect(parseWorkbook("<workbook><sheets/></workbook>").date1904).toBe(false);
		expect(
			parseWorkbook('<workbook><workbookPr date1904="1"/><sheets/></workbook>').date1904,
		).toBe(true);
		expect(
			parseWorkbook('<workbook><workbookPr date1904="true"/><sheets/></workbook>').date1904,
		).toBe(true);
		expect(
			parseWorkbook('<workbook><workbookPr date1904="0"/><sheets/></workbook>').date1904,
		).toBe(false);
	});
});

describe("parseWorkbook — defined names (F10.1)", () => {
	// One sheet, so localSheetId 0 is in range and 5 is out of range.
	const wrap = (definedNames: string) =>
		`<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets>` +
		`<definedNames>${definedNames}</definedNames></workbook>`;

	it("reads a global name and a sheet-local hidden name, in document order", () => {
		const xml = wrap(
			'<definedName name="Total">Sheet1!$A$1:$A$9</definedName>' +
				'<definedName name="Local" localSheetId="0" hidden="1">Sheet1!$B$2</definedName>',
		);
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "Total", refersTo: "Sheet1!$A$1:$A$9" },
			{ name: "Local", refersTo: "Sheet1!$B$2", localSheetId: 0, hidden: true },
		]);
	});

	it("keeps a spec built-in like _xlnm.Print_Area", () => {
		const xml = wrap(
			'<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$3</definedName>',
		);
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "_xlnm.Print_Area", refersTo: "Sheet1!$A$1:$C$3", localSheetId: 0 },
		]);
	});

	it("DROPS a name the strict writer could not re-emit (F10.1 decision 3, named degradation)", () => {
		const xml = wrap(
			'<definedName name="Good">Sheet1!$A$1</definedName>' +
				'<definedName name="Bad Name">Sheet1!$A$1</definedName>' + // whitespace in the name
				'<definedName name="A1">Sheet1!$A$1</definedName>' + // cell-reference shape
				'<definedName name="_xlnm.Nope">Sheet1!$A$1</definedName>' + // reserved prefix, not a built-in
				'<definedName name="Empty"></definedName>' + // empty refersTo
				'<definedName name="OffSheet" localSheetId="5">Sheet1!$A$1</definedName>', // scope past the sheet list
		);
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "Good", refersTo: "Sheet1!$A$1" },
		]);
	});

	it("decodes an ST_Xstring @name — a stored _x005F_x0041_ IS the name _x0041_ (F9.6 parity)", () => {
		const xml = wrap('<definedName name="_x005F_x0041_">Sheet1!$A$1</definedName>');
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "_x0041_", refersTo: "Sheet1!$A$1" },
		]);
	});

	it("drops a refersTo that decodes to an XML-illegal control character", () => {
		// A numeric char ref the tokenizer decodes to U+0001 — cannot re-emit as element text.
		const xml = wrap('<definedName name="Ctrl">Sheet1!$A$1&#1;</definedName>');
		expect(parseWorkbook(xml).definedNames).toEqual([]);
	});

	it("is empty when the workbook declares no names", () => {
		expect(
			parseWorkbook('<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>')
				.definedNames,
		).toEqual([]);
	});
});
