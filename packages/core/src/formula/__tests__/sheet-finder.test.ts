import { describe, expect, it, onTestFinished, vi } from "vitest";
import { rangeSheet, sheetFinder } from "../eval";

describe("sheetFinder", () => {
	it("finds a sheet spelled exactly as the qualifier", () => {
		expect(sheetFinder(["Data", "Main"])("Data")).toBe("Data");
		expect(sheetFinder(["Data", "Main"])("Main")).toBe("Main");
	});

	it("finds a sheet whose name differs from the qualifier only in case", () => {
		const find = sheetFinder(["Data", "My Sheet"]);
		expect(find("data")).toBe("Data");
		expect(find("DATA")).toBe("Data");
		expect(find("my SHEET")).toBe("My Sheet");
	});

	it("folds case as the reader does for duplicate names, beyond ASCII", () => {
		expect(sheetFinder(["Ärger"])("ärger")).toBe("Ärger");
		expect(sheetFinder(["ΣΥΝΟΛΟ"])("σύνολο")).toBeUndefined(); // an accent is not case
		expect(sheetFinder(["ΣΥΝΟΛΟ"])("συνολο")).toBe("ΣΥΝΟΛΟ");
		expect(sheetFinder(["Straße"])("STRASSE")).toBeUndefined(); // the reader keeps both names
		expect(sheetFinder(["Straße"])("straße")).toBe("Straße");
	});

	it("prefers the exact spelling, then the first sheet in tab order", () => {
		const find = sheetFinder(["data", "Data"]);
		expect(find("Data")).toBe("Data");
		expect(find("data")).toBe("data");
		expect(find("DATA")).toBe("data");
	});

	it("skips a sheet name that is not a string", () => {
		const find = sheetFinder([5 as never, null as never, "Data"]);
		expect(find("data")).toBe("Data");
		expect(find("5")).toBeUndefined();
	});

	it("folds each qualifier once, however often it is looked up", () => {
		const find = sheetFinder(["Data"]);
		const lower = vi.spyOn(String.prototype, "toLowerCase");
		onTestFinished(() => lower.mockRestore());
		const long = "x".repeat(8000);
		for (let i = 0; i < 1000; i++) {
			expect(find(long)).toBeUndefined();
			expect(find("DATA")).toBe("Data");
		}
		expect(lower).toHaveBeenCalledTimes(2);
	});

	it("finds nothing for a qualifier that names no sheet", () => {
		expect(sheetFinder(["Data"])("Nope")).toBeUndefined();
		expect(sheetFinder(["Data"])("Data ")).toBeUndefined();
		expect(sheetFinder([])("Data")).toBeUndefined();
	});
});

describe("rangeSheet", () => {
	const find = sheetFinder(["Data", "Main"]);

	it("takes the sheet of the qualified end, in any case", () => {
		expect(rangeSheet("Data", undefined, "Main", find)).toBe("Data");
		expect(rangeSheet(undefined, "data", "Main", find)).toBe("Data");
		expect(rangeSheet("DATA", "data", "Main", find)).toBe("Data");
	});

	it("takes the left end's sheet when the two ends name different sheets", () => {
		expect(rangeSheet("Main", "Data", "Main", find)).toBe("Main");
	});

	it("finds no sheet when either qualifier names none", () => {
		expect(rangeSheet("Nope", undefined, "Main", find)).toBeUndefined();
		expect(rangeSheet("Data", "Nope", "Main", find)).toBeUndefined();
		expect(rangeSheet(undefined, "Nope", "Main", find)).toBeUndefined();
	});

	it("falls back to the formula's own sheet when neither end is qualified", () => {
		expect(rangeSheet(undefined, undefined, "Main", find)).toBe("Main");
	});
});
