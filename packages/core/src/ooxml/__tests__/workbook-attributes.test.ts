import { describe, expect, it } from "vitest";
import { booleanAttribute, workbookProtectionFromAttributes } from "../workbook-attributes";

describe("workbook protection attributes", () => {
	it("parses only the four OOXML boolean spellings", () => {
		expect(booleanAttribute("1")).toBe(true);
		expect(booleanAttribute("true")).toBe(true);
		expect(booleanAttribute("0")).toBe(false);
		expect(booleanAttribute("false")).toBe(false);
		expect(booleanAttribute("TRUE")).toBeUndefined();
		expect(booleanAttribute(undefined)).toBeUndefined();
	});

	it("parses every OOXML boolean spelling and preserves explicit false", () => {
		expect(
			workbookProtectionFromAttributes({ lockStructure: "1", lockWindows: "false" }),
		).toEqual({ lockStructure: true, lockWindows: false });
		expect(
			workbookProtectionFromAttributes({ lockStructure: "true", lockWindows: "0" }),
		).toEqual({ lockStructure: true, lockWindows: false });
	});

	it("carries password and hash fields verbatim", () => {
		expect(
			workbookProtectionFromAttributes({
				workbookPassword: "C258",
				workbookAlgorithmName: "SHA-512",
				workbookHashValue: "hash==",
				workbookSaltValue: "salt==",
			}),
		).toEqual({
			workbookPassword: "C258",
			workbookAlgorithmName: "SHA-512",
			workbookHashValue: "hash==",
			workbookSaltValue: "salt==",
		});
	});

	it("accepts both unsigned-integer spin-count boundaries", () => {
		expect(workbookProtectionFromAttributes({ workbookSpinCount: "0" })).toEqual({
			workbookSpinCount: 0,
		});
		expect(workbookProtectionFromAttributes({ workbookSpinCount: "4294967295" })).toEqual({
			workbookSpinCount: 4_294_967_295,
		});
	});

	it.each([
		"",
		"-1",
		"1.5",
		"1e3",
		" 1",
		"+1",
		"4294967296",
	])("drops malformed or out-of-range spin count %j", (workbookSpinCount) => {
		expect(workbookProtectionFromAttributes({ workbookSpinCount })).toBeUndefined();
	});

	it("drops malformed booleans while keeping other valid fields", () => {
		expect(
			workbookProtectionFromAttributes({
				lockStructure: "TRUE",
				lockWindows: "yes",
				workbookPassword: "ABCD",
			}),
		).toEqual({ workbookPassword: "ABCD" });
	});

	it("returns undefined for empty or wholly malformed attributes", () => {
		expect(workbookProtectionFromAttributes({})).toBeUndefined();
		expect(
			workbookProtectionFromAttributes({ lockStructure: "yes", workbookSpinCount: "NaN" }),
		).toBeUndefined();
	});

	it("preserves explicitly present empty string fields", () => {
		expect(workbookProtectionFromAttributes({ workbookHashValue: "" })).toEqual({
			workbookHashValue: "",
		});
	});
});
