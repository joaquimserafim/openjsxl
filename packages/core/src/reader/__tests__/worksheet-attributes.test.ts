import { describe, expect, it } from "vitest";
import {
	booleanAttribute,
	clampedDoubleAttribute,
	enumAttribute,
	headerFooterFromAttributes,
	isHeaderFooterChild,
	pageMarginsFromAttributes,
	pageSetupFromAttributes,
	printOptionsFromAttributes,
	sheetProtectionFromAttributes,
	trueAttribute,
	unsignedIntAttribute,
} from "../worksheet-attributes";

describe("worksheet attribute helpers", () => {
	it("parses OOXML boolean spellings without accepting look-alikes", () => {
		expect(booleanAttribute("1")).toBe(true);
		expect(booleanAttribute("true")).toBe(true);
		expect(booleanAttribute("0")).toBe(false);
		expect(booleanAttribute("false")).toBe(false);
		expect(booleanAttribute("TRUE")).toBeUndefined();
		expect(booleanAttribute(undefined)).toBeUndefined();
		expect(trueAttribute("1")).toBe(true);
		expect(trueAttribute("true")).toBe(true);
		expect(trueAttribute("0")).toBe(false);
		expect(trueAttribute("invalid")).toBe(false);
	});

	it("clamps finite doubles and drops non-finite values", () => {
		expect(clampedDoubleAttribute("1.25", 49)).toBe(1.25);
		expect(clampedDoubleAttribute("-2", 49)).toBe(0);
		expect(clampedDoubleAttribute("99", 49)).toBe(49);
		expect(clampedDoubleAttribute("", 49)).toBe(0); // existing tolerant Number("") behavior
		expect(clampedDoubleAttribute("NaN", 49)).toBeUndefined();
		expect(clampedDoubleAttribute("Infinity", 49)).toBeUndefined();
		expect(clampedDoubleAttribute(undefined, 49)).toBeUndefined();
	});

	it("accepts only canonical bounded unsigned integers", () => {
		expect(unsignedIntAttribute("0")).toBe(0);
		expect(unsignedIntAttribute("4294967295")).toBe(4_294_967_295);
		expect(unsignedIntAttribute("4294967296")).toBeUndefined();
		expect(unsignedIntAttribute("11", 10)).toBeUndefined();
		for (const invalid of ["", "-1", "1.5", "1e3", " 1"]) {
			expect(unsignedIntAttribute(invalid)).toBeUndefined();
		}
	});

	it("keeps only allowed enum values", () => {
		const allowed = ["portrait", "landscape"] as const;
		expect(enumAttribute("landscape", allowed)).toBe("landscape");
		expect(enumAttribute("sideways", allowed)).toBeUndefined();
		expect(enumAttribute(undefined, allowed)).toBeUndefined();
	});

	it("builds sheet protection while dropping invalid spin counts", () => {
		expect(
			sheetProtectionFromAttributes({
				sheet: "1",
				formatCells: "0",
				password: "C258",
				algorithmName: "SHA-512",
				hashValue: "h==",
				saltValue: "s==",
				spinCount: "4294967295",
			}),
		).toEqual({
			sheet: true,
			formatCells: false,
			password: "C258",
			algorithmName: "SHA-512",
			hashValue: "h==",
			saltValue: "s==",
			spinCount: 4_294_967_295,
		});
		expect(
			sheetProtectionFromAttributes({ sheet: "1", spinCount: "999999999999999999999" }),
		).toEqual({ sheet: true });
		expect(sheetProtectionFromAttributes({ sheet: "yes" })).toBeUndefined();
		expect(sheetProtectionFromAttributes({})).toBeUndefined();
	});

	it("builds print options and preserves explicit false values", () => {
		expect(
			printOptionsFromAttributes({
				gridLines: "1",
				headings: "0",
				horizontalCentered: "true",
				verticalCentered: "false",
			}),
		).toEqual({
			gridLines: true,
			headings: false,
			horizontalCentered: true,
			verticalCentered: false,
		});
		expect(printOptionsFromAttributes({ gridLines: "yes" })).toBeUndefined();
		expect(printOptionsFromAttributes({})).toBeUndefined();
	});

	it("requires all six margins and clamps the complete record", () => {
		expect(
			pageMarginsFromAttributes({
				left: "-1",
				right: "99",
				top: "1",
				bottom: "1",
				header: "0.3",
				footer: "0.3",
			}),
		).toEqual({ left: 0, right: 49, top: 1, bottom: 1, header: 0.3, footer: 0.3 });
		expect(
			pageMarginsFromAttributes({
				left: "1",
				right: "1",
				top: "1",
				bottom: "1",
				header: "1",
			}),
		).toBeUndefined();
	});

	it("builds every supported page-setup field and clamps scale", () => {
		expect(
			pageSetupFromAttributes({
				paperSize: "9",
				orientation: "landscape",
				scale: "5",
				fitToWidth: "1",
				fitToHeight: "0",
				firstPageNumber: "2",
				useFirstPageNumber: "1",
				blackAndWhite: "0",
				draft: "true",
				cellComments: "atEnd",
				pageOrder: "overThenDown",
			}),
		).toEqual({
			paperSize: 9,
			orientation: "landscape",
			scale: 10,
			fitToWidth: 1,
			fitToHeight: 0,
			firstPageNumber: 2,
			useFirstPageNumber: true,
			blackAndWhite: false,
			draft: true,
			cellComments: "atEnd",
			pageOrder: "overThenDown",
		});
		expect(pageSetupFromAttributes({ scale: "900" })).toEqual({ scale: 400 });
		expect(
			pageSetupFromAttributes({ orientation: "sideways", paperSize: "1e3", draft: "yes" }),
		).toBeUndefined();
	});

	it("builds header/footer flags and recognizes text children", () => {
		expect(
			headerFooterFromAttributes({
				differentOddEven: "1",
				differentFirst: "0",
				scaleWithDoc: "true",
				alignWithMargins: "false",
			}),
		).toEqual({
			differentOddEven: true,
			differentFirst: false,
			scaleWithDoc: true,
			alignWithMargins: false,
		});
		expect(headerFooterFromAttributes({ differentFirst: "yes" })).toEqual({});
		expect(isHeaderFooterChild("oddHeader")).toBe(true);
		expect(isHeaderFooterChild("firstFooter")).toBe(true);
		expect(isHeaderFooterChild("headerFooter")).toBe(false);
	});
});
