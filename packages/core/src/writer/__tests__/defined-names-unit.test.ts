import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import type { DefinedNameProblem } from "../../ooxml/name";
import {
	definedNameProblemMessage,
	definedNamesXml,
	filterDatabaseName,
	qualifiedAbsoluteRange,
	validateDefinedNames,
} from "../defined-names";

describe("definedNamesXml", () => {
	it("omits the container for empty input", () => {
		// Input `[]` -> no workbook.xml fragment, preserving the names-free byte path.
		expect(definedNamesXml([])).toBe("");
	});

	it("protects escape-shaped logical text and emits scoped attributes in order", () => {
		// Logical `_x0041_` -> wire `_x005F_x0041_`, in both the name and complete formula.
		expect(
			definedNamesXml([
				{
					name: "_x0041_",
					refersTo: "'_x0041_'!B4",
					localSheetId: 0,
					hidden: true,
				},
			]),
		).toBe(
			'<definedNames><definedName name="_x005F_x0041_" localSheetId="0" hidden="1">\'_x005F_x0041_\'!B4</definedName></definedNames>',
		);
	});

	it("XML-escapes after applying ST_Xstring protection", () => {
		// Logical formula text `"_x0041_<&"` -> protected escape text plus safe XML entities.
		expect(definedNamesXml([{ name: "Greeting", refersTo: '"_x0041_<&"' }])).toBe(
			'<definedNames><definedName name="Greeting">"_x005F_x0041_&lt;&amp;"</definedName></definedNames>',
		);
	});
});

describe("validateDefinedNames", () => {
	it("turns absent or empty input into an empty trusted array", () => {
		// Inputs `undefined` and `[]` -> `[]`.
		expect(validateDefinedNames(undefined, 1)).toEqual([]);
		expect(validateDefinedNames([], 1)).toEqual([]);
	});

	it("reads each property once and drops explicit hidden false from trusted output", () => {
		const reads = { name: 0, refersTo: 0, localSheetId: 0, hidden: 0 };
		const entry = Object.defineProperties(
			{},
			{
				name: {
					enumerable: true,
					get: () => {
						reads.name++;
						return "Scoped";
					},
				},
				refersTo: {
					enumerable: true,
					get: () => {
						reads.refersTo++;
						return "Data!$A$1";
					},
				},
				localSheetId: {
					enumerable: true,
					get: () => {
						reads.localSheetId++;
						return 0;
					},
				},
				hidden: {
					enumerable: true,
					get: () => {
						reads.hidden++;
						return false;
					},
				},
			},
		);

		// External input -> one stable logical object; defaults are omitted from trusted output.
		expect(validateDefinedNames([entry], 1)).toEqual([
			{ name: "Scoped", refersTo: "Data!$A$1", localSheetId: 0 },
		]);
		expect(reads).toEqual({ name: 1, refersTo: 1, localSheetId: 1, hidden: 1 });
	});

	it("keeps equal names in different scopes but rejects a same-scope duplicate", () => {
		expect(
			validateDefinedNames(
				[
					{ name: "Rate", refersTo: "1" },
					{ name: "Rate", refersTo: "2", localSheetId: 0 },
				],
				1,
			),
		).toHaveLength(2);
		expect(() =>
			validateDefinedNames(
				[
					{ name: "Rate", refersTo: "1" },
					{ name: "rate", refersTo: "2" },
				],
				1,
			),
		).toThrowError(/duplicate name/);
	});

	it("returns typed validation errors with a specific reason", () => {
		const invalid = [{ name: "A1", refersTo: "1" }];
		expect(() => validateDefinedNames(invalid, 1)).toThrowError(XlsxError);
		expect(() => validateDefinedNames(invalid, 1)).toThrowError(/cell reference/);
		expect(() =>
			validateDefinedNames([{ name: "_xlnm._FilterDatabase", refersTo: "Data!$A$1" }], 1),
		).toThrowError(/SheetInput\.autoFilter/);
	});
});

describe("defined-name helper messages and references", () => {
	it.each<[DefinedNameProblem, string]>([
		["empty", "must not be empty"],
		["too-long", "exceeds"],
		["not-xml-safe", "not allowed in XML"],
		["whitespace", "must not contain whitespace"],
		["bad-start", "must start with"],
		["cell-ref", "must not look like a cell reference"],
		["bad-builtin", "reserved"],
	])("maps %s to a specific explanation", (problem, expected) => {
		expect(definedNameProblemMessage("Example", problem)).toContain(expected);
	});

	it("quotes sheet names and makes canonical cell and range references absolute", () => {
		expect(qualifiedAbsoluteRange("Bob's", "A1:C3")).toBe("'Bob''s'!$A$1:$C$3");
		expect(qualifiedAbsoluteRange("Data", "B4")).toBe("'Data'!$B$4");
	});

	it("keeps a non-canonical reference unchanged after qualifying its sheet", () => {
		expect(qualifiedAbsoluteRange("Data", "A:A")).toBe("'Data'!A:A");
	});
});

describe("filterDatabaseName", () => {
	it("quotes apostrophes and makes a canonical range absolute", () => {
		// Input `("Bob's", 2, "A1:C3")` -> hidden sheet-local absolute range.
		expect(filterDatabaseName("Bob's", 2, "A1:C3")).toEqual({
			name: "_xlnm._FilterDatabase",
			refersTo: "'Bob''s'!$A$1:$C$3",
			localSheetId: 2,
			hidden: true,
		});
	});
});
