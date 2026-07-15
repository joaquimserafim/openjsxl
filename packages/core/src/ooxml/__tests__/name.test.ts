import { describe, expect, it } from "vitest";
import { MAX_FORMULA_LEN } from "../formula";
import {
	definedNameEmittable,
	definedNameProblem,
	MAX_NAME_LEN,
	nameProblem,
	normalizeName,
} from "../name";

// The shared identifier core (nameProblem/normalizeName) is exercised in depth via its table-name
// aliases in table.test.ts; these focus on the F10.1 defined-name layer plus a couple of core spot
// checks to pin that the extraction preserved behavior.

describe("nameProblem — shared identifier core (post-extraction spot checks)", () => {
	it("accepts a legal identifier and names the first broken rule otherwise", () => {
		expect(nameProblem("Sales")).toBeUndefined();
		expect(nameProblem("_Total")).toBeUndefined();
		expect(nameProblem("データ")).toBeUndefined(); // any Unicode letter may start a name (F9.6)
		expect(nameProblem("")).toBe("empty");
		expect(nameProblem("a".repeat(MAX_NAME_LEN + 1))).toBe("too-long");
		expect(nameProblem("a\x01b")).toBe("not-xml-safe");
		expect(nameProblem("My Name")).toBe("whitespace");
		expect(nameProblem("1abc")).toBe("bad-start");
		expect(nameProblem("A1")).toBe("cell-ref");
		expect(nameProblem("R")).toBe("cell-ref");
	});

	it("normalizeName leaves a legal name untouched (byte-identity) and repairs an illegal one", () => {
		expect(normalizeName("Sales")).toBe("Sales");
		expect(nameProblem(normalizeName("My Bad Name"))).toBeUndefined();
		expect(nameProblem(normalizeName("1st Quarter"))).toBeUndefined();
	});
});

describe("definedNameProblem — defined-name grammar + the reserved _xlnm. prefix", () => {
	it("accepts an ordinary legal name", () => {
		expect(definedNameProblem("Sales")).toBeUndefined();
		expect(definedNameProblem("_Total")).toBeUndefined();
		expect(definedNameProblem("Q1.Revenue")).toBeUndefined(); // periods are allowed in the tail
	});

	it("accepts the reserved _xlnm. prefix ONLY for a spec built-in suffix", () => {
		expect(definedNameProblem("_xlnm.Print_Area")).toBeUndefined();
		expect(definedNameProblem("_xlnm.Print_Titles")).toBeUndefined();
		expect(definedNameProblem("_xlnm._FilterDatabase")).toBeUndefined();
		expect(definedNameProblem("_xlnm.Database")).toBeUndefined();
		// A non-built-in suffix is rejected — the namespace is reserved.
		expect(definedNameProblem("_xlnm.MyArea")).toBe("bad-builtin");
		expect(definedNameProblem("_xlnm.")).toBe("bad-builtin");
	});

	it("matches the built-in suffix and the prefix case-insensitively (a foreign producer may vary casing)", () => {
		expect(definedNameProblem("_xlnm.print_area")).toBeUndefined();
		expect(definedNameProblem("_XLNM.Print_Area")).toBeUndefined();
	});

	it("still rejects every ordinary illegality (whitespace, cell-ref shape, bad start, …)", () => {
		expect(definedNameProblem("My Name")).toBe("whitespace");
		expect(definedNameProblem("A1")).toBe("cell-ref");
		expect(definedNameProblem("1abc")).toBe("bad-start");
		expect(definedNameProblem("")).toBe("empty");
	});

	it("treats a name that merely starts with _xlnm (no dot) as an ordinary name", () => {
		expect(definedNameProblem("_xlnmFoo")).toBeUndefined(); // not the reserved `_xlnm.` prefix
	});
});

describe("definedNameEmittable — the reader's drop-to-writer-legal predicate", () => {
	const ok = { name: "Sales", refersTo: "'Sheet1'!$A$1:$B$2" };

	it("keeps a name the strict writer could re-emit", () => {
		expect(definedNameEmittable(ok, 1)).toBe(true);
		expect(definedNameEmittable({ ...ok, localSheetId: 0 }, 2)).toBe(true);
		expect(
			definedNameEmittable({ name: "_xlnm.Print_Area", refersTo: "Sheet1!$1:$1" }, 1),
		).toBe(true);
	});

	it("drops an illegal name", () => {
		expect(definedNameEmittable({ name: "Bad Name", refersTo: "1" }, 1)).toBe(false);
		expect(definedNameEmittable({ name: "A1", refersTo: "1" }, 1)).toBe(false);
		expect(definedNameEmittable({ name: "_xlnm.Nope", refersTo: "1" }, 1)).toBe(false);
	});

	it("drops an empty, oversized, =-prefixed, or XML-unsafe refersTo", () => {
		expect(definedNameEmittable({ name: "X", refersTo: "" }, 1)).toBe(false);
		expect(
			definedNameEmittable({ name: "X", refersTo: "a".repeat(MAX_FORMULA_LEN + 1) }, 1),
		).toBe(false);
		expect(definedNameEmittable({ name: "X", refersTo: "=A1" }, 1)).toBe(false);
		expect(definedNameEmittable({ name: "X", refersTo: "A1\x01" }, 1)).toBe(false); // e.g. injected via &#1;
	});

	it("drops a sheet-scope pointing past the sheet list", () => {
		expect(definedNameEmittable({ ...ok, localSheetId: 5 }, 2)).toBe(false);
		expect(definedNameEmittable({ ...ok, localSheetId: -1 }, 2)).toBe(false);
	});
});
