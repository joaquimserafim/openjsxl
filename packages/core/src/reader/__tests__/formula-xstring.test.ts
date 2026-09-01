import { describe, expect, it } from "vitest";
import { MAX_FORMULA_LEN } from "../../ooxml/formula";
import { encodeXstring } from "../../ooxml/xstring";
import { parseFormulas } from "../worksheet";

// F11.20 — ST_Formula is a restriction of ST_Xstring, so `<f>` text is decoded exactly once, on the
// COMPLETE accumulated element text, BEFORE the length ceiling, master registration, and dependent
// translation. A stored `_x005F_x0041_` is the literal `_x0041_`; a raw `_x0041_` is `A` (as Excel
// reads it). A decoded value the strict writer cannot emit degrades to "no formula".

const sheet = (body: string) =>
	`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
const cell = (f: string) => sheet(`<row r="1"><c r="A1"><f>${f}</f><v>1</v></c></row>`);

describe("parseFormulas — ST_Xstring decoding (F11.20)", () => {
	it("decodes a raw escape to its character and a protected escape to the literal", () => {
		expect(parseFormulas(cell("'_x0041_'!B1")).get("A1")).toBe("'A'!B1");
		expect(parseFormulas(cell("'_x005F_x0041_'!B1")).get("A1")).toBe("'_x0041_'!B1");
		expect(parseFormulas(cell('"_x0041_"&amp;"_x005F_x0041_"')).get("A1")).toBe(
			'"A"&"_x0041_"',
		);
	});

	it("accepts lowercase hex, never double-decodes, and leaves an uppercase X alone", () => {
		expect(parseFormulas(cell("_x005f_x0041_")).get("A1")).toBe("_x0041_"); // lowercase hex
		// One non-reentrant pass: whatever the writer's encoder protects decodes back to exactly the
		// logical text — never a second decode of the protected output.
		const logical = "_x005F_x0041_";
		expect(parseFormulas(cell(encodeXstring(logical))).get("A1")).toBe(logical);
		expect(parseFormulas(cell("_X0041_")).get("A1")).toBe("_X0041_"); // not an escape
	});

	it("decodes the COMPLETE text — an escape split by an entity or CDATA still decodes", () => {
		// The tokenizer emits separate text events around an entity / CDATA section; the decode must
		// run once over the accumulated element text, not per event.
		expect(parseFormulas(cell("_x00&#52;1_")).get("A1")).toBe("A");
		expect(parseFormulas(cell("<![CDATA[_x00]]>41_")).get("A1")).toBe("A");
	});

	it("decodes a shared master BEFORE translating its dependents", () => {
		const f = parseFormulas(
			sheet(
				'<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">\'_x0041_\'!B1+_x005F_x0041_</f><v>1</v></c></row>' +
					'<row r="2"><c r="A2"><f t="shared" si="0"/><v>1</v></c></row>',
			),
		);
		expect(f.get("A1")).toBe("'A'!B1+_x0041_"); // master: decoded, verbatim
		expect(f.get("A2")).toBe("'A'!B2+_x0041_"); // dependent: decoded text, then row-shifted
	});

	it("applies the master length ceiling to the LOGICAL text, not the wider wire text", () => {
		// 1200 raw escapes: wire 8,400 chars (> MAX_FORMULA_LEN), logical 1,200 chars (≤ MAX_FORMULA_LEN).
		const wire = "_x0041_".repeat(1200);
		expect(wire.length).toBeGreaterThan(MAX_FORMULA_LEN);
		const f = parseFormulas(
			sheet(
				`<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">${wire}</f><v>1</v></c></row>` +
					'<row r="2"><c r="A2"><f t="shared" si="0"/><v>1</v></c></row>',
			),
		);
		expect(f.get("A1")).toBe("A".repeat(1200));
		expect(f.get("A2")).toBe("A".repeat(1200)); // master was registered → dependent resolved
	});

	it("degrades a formula whose DECODED text the writer would refuse (control char / lone surrogate)", () => {
		expect(parseFormulas(cell("A1_x0001_")).has("A1")).toBe(false);
		expect(parseFormulas(cell("A1_xD800_")).has("A1")).toBe(false);
		// An unsafe shared master registers nothing — its dependents get no formula either.
		const f = parseFormulas(
			sheet(
				'<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1_x0000_</f><v>1</v></c></row>' +
					'<row r="2"><c r="A2"><f t="shared" si="0"/><v>1</v></c></row>',
			),
		);
		expect(f.size).toBe(0);
	});

	it("strips one leading = after decoding — raw or escaped — so the bridge can re-emit it (stored form)", () => {
		expect(parseFormulas(cell("=A1")).get("A1")).toBe("A1");
		expect(parseFormulas(cell("_x003D_A1")).get("A1")).toBe("A1");
		expect(parseFormulas(cell("_x003D_")).has("A1")).toBe(false); // nothing usable remains
	});

	it("covers the reference forms the escape can hide in (Task D)", () => {
		expect(parseFormulas(cell("_x0041_!A1")).get("A1")).toBe("A!A1"); // unquoted sheet ref
		expect(parseFormulas(cell("'_x0041_:_x0042_'!A1")).get("A1")).toBe("'A:B'!A1"); // 3-D ref
		expect(parseFormulas(cell("'it''s_x0041_'!A1")).get("A1")).toBe("'it''sA'!A1"); // doubled '
		expect(parseFormulas(cell("[1]_x0041_!A1")).get("A1")).toBe("[1]A!A1"); // external text
		expect(parseFormulas(cell("_x0041_name*2")).get("A1")).toBe("Aname*2"); // defined-name ref
		// Overlapping protection: whatever the encoder emits for a logical value whose escape shapes
		// overlap decodes back to exactly that value.
		const overlap = "_x0041_x0042_";
		expect(parseFormulas(cell(encodeXstring(overlap))).get("A1")).toBe(overlap);
		// Array master: same decode as a plain formula.
		const array = sheet(
			'<row r="1"><c r="A1"><f t="array" ref="A1:A3">_x0041_1:A3*2</f><v>2</v></c></row>',
		);
		expect(parseFormulas(array).get("A1")).toBe("A1:A3*2");
	});

	it("keeps clean formula text verbatim (decode is the identity when nothing is escaped)", () => {
		expect(parseFormulas(cell("SUM(A1:A9)*2")).get("A1")).toBe("SUM(A1:A9)*2");
		expect(parseFormulas(cell("A1&lt;B1")).get("A1")).toBe("A1<B1");
		expect(parseFormulas(cell("a_x_b")).get("A1")).toBe("a_x_b"); // escape-SHAPED but not an escape
	});
});
