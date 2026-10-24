import { describe, expect, it } from "vitest";
import { claimPart } from "../part-claims";

describe("claimPart", () => {
	it("claims a path once per set", () => {
		const claimed = new Set<string>();
		expect(claimPart(claimed, "xl/worksheets/sheet1.xml")).toBe(true);
		expect(claimPart(claimed, "xl/worksheets/sheet1.xml")).toBe(false);
		expect(claimPart(claimed, "xl/worksheets/sheet2.xml")).toBe(true);
		expect(claimPart(new Set(), "xl/worksheets/sheet1.xml")).toBe(true);
		expect([...claimed]).toEqual(["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]);
	});
});
