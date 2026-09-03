import { describe, expect, it } from "vitest";
import { MAX_FORMULA_LEN } from "../formula";
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

	it("decodes each sheet name as one ST_Xstring attribute value", () => {
		const xml =
			"<workbook><sheets>" +
			'<sheet name="Plain" r:id="rId1"/>' +
			'<sheet name="_x0041_" r:id="rId2"/>' +
			'<sheet name="_x005F_x0041_" r:id="rId3"/>' +
			'<sheet name="_x005f_x0042_" r:id="rId4"/>' +
			'<sheet name="&#95;x0043_ &amp; Co" r:id="rId5"/>' +
			"</sheets></workbook>";
		expect(parseWorkbook(xml).sheets.map((sheet) => sheet.name)).toEqual([
			"Plain",
			"A",
			"_x0041_",
			"_x0042_",
			"C & Co",
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

	it("decodes the complete ST_Formula value once, without interpreting formula tokens", () => {
		// CT_DefinedName extends ST_Formula, and ST_Formula restricts ST_Xstring. The whole text is
		// decoded uniformly: quoted/bare sheet names, strings, names, structured refs, and externals.
		const protectedFormula =
			"SUM('_x005F_x0041_:_x005F_x0042_'!A1)+_x005F_x0041_+" +
			"Table1[_x005F_x0041_]+\"_x005F_x0041_\"+'[1]_x005F_x0041_'!A1";
		const logicalFormula =
			"SUM('_x0041_:_x0042_'!A1)+_x0041_+" + "Table1[_x0041_]+\"_x0041_\"+'[1]_x0041_'!A1";
		const xml = wrap(
			`<definedName name="Protected">${protectedFormula}</definedName>` +
				"<definedName name=\"Raw\">'_x0041_'!B4</definedName>" +
				'<definedName name="Lower">"_x005f_x0041_"</definedName>' +
				'<definedName name="Overlap">"_x005F_x005F_x005F_x0041_"</definedName>',
		);
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "Protected", refersTo: logicalFormula },
			// A raw escape is data on the wire, so decoding does not depend on whether sheet A exists.
			{ name: "Raw", refersTo: "'A'!B4" },
			{ name: "Lower", refersTo: '"_x0041_"' },
			{ name: "Overlap", refersTo: '"_x005F_x0041_"' },
		]);
	});

	it("accumulates split text and CDATA before decoding the complete formula", () => {
		// Producers may split one escape sequence across tokenizer events. Decoding each chunk would
		// miss it, so the parser must join all text first and perform one non-reentrant decode pass.
		const xml = wrap(
			"<definedName name=\"Split\">'_x005F_<![CDATA[x0041_]]>'!B4</definedName>",
		);
		expect(parseWorkbook(xml).definedNames).toEqual([
			{ name: "Split", refersTo: "'_x0041_'!B4" },
		]);
	});

	it("validates the decoded logical formula length, not its longer protected wire form", () => {
		// 1,171 protected underscores occupy 8,197 wire characters but only 1,171 logical ones.
		const wire = "_x005F_".repeat(Math.floor(MAX_FORMULA_LEN / 7) + 1);
		expect(wire.length).toBeGreaterThan(MAX_FORMULA_LEN);
		const [name] = parseWorkbook(
			wrap(`<definedName name="LongWire">${wire}</definedName>`),
		).definedNames;
		expect(name?.refersTo).toBe("_".repeat(wire.length / 7));
	});

	it("applies stored-form validation after ST_Xstring decoding", () => {
		// `_x003D_` is `=`. Accepting this raw wire text before decoding would leak a logical formula
		// that violates the public no-leading-equals invariant.
		const xml = wrap('<definedName name="EncodedEquals">_x003D_A1</definedName>');
		expect(parseWorkbook(xml).definedNames).toEqual([]);
	});

	it.each([
		"&#1;",
		"_x0001_",
	])("drops a refersTo that decodes %s to an XML-illegal control character", (escapedControl) => {
		// Validate after both XML-entity and ST_Xstring decoding so the writer can re-emit every
		// surviving logical value safely.
		const xml = wrap(`<definedName name="Ctrl">Sheet1!$A$1${escapedControl}</definedName>`);
		expect(parseWorkbook(xml).definedNames).toEqual([]);
	});

	it("is empty when the workbook declares no names", () => {
		expect(
			parseWorkbook('<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>')
				.definedNames,
		).toEqual([]);
	});
});

describe("parseWorkbook — <workbookPr date1904> (F11.10)", () => {
	const wb = (attr: string) =>
		`<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr${attr}/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;

	it("reads the xsd:boolean lexical forms, with whitespace collapsed, case-sensitively", () => {
		expect(parseWorkbook(wb(' date1904="1"')).date1904).toBe(true);
		expect(parseWorkbook(wb(' date1904="true"')).date1904).toBe(true);
		expect(parseWorkbook(wb(' date1904=" 1 "')).date1904).toBe(true); // whiteSpace=collapse
		expect(parseWorkbook(wb(' date1904="0"')).date1904).toBe(false);
		expect(parseWorkbook(wb(' date1904="false"')).date1904).toBe(false);
		expect(parseWorkbook(wb(' date1904="TRUE"')).date1904).toBe(false); // schema-invalid → default
		expect(parseWorkbook(wb("")).date1904).toBe(false); // absent → the 1900 default
	});
});
