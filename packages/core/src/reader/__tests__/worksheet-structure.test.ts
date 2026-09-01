import { describe, expect, it } from "vitest";
import {
	createWorksheetStructure,
	isWorksheetRegion,
	pushWorksheetClose,
	pushWorksheetOpen,
} from "../worksheet-structure";

// These tests drive the constant-space container tracker directly. It decides whether a style
// token belongs to worksheet/cols or worksheet/sheetData; cell and style parsing stay elsewhere.

describe("worksheet structural scope", () => {
	it("opens only prefixed canonical containers that are direct worksheet children", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "x:worksheet", false);
		pushWorksheetOpen(state, "x:cols", false);
		expect(isWorksheetRegion(state, "cols")).toBe(true);
		expect(pushWorksheetClose(state, "x:cols")).toBe("cols");
		pushWorksheetOpen(state, "x:sheetData", false);
		expect(isWorksheetRegion(state, "sheetData")).toBe(true);
		expect(pushWorksheetClose(state, "x:sheetData")).toBe("sheetData");
		pushWorksheetClose(state, "x:worksheet");
		expect(state.phase).toBe("closed");
	});

	it("does not activate self-closing containers or a non-worksheet document root", () => {
		const selfClosing = createWorksheetStructure();
		pushWorksheetOpen(selfClosing, "worksheet", false);
		pushWorksheetOpen(selfClosing, "cols", true);
		expect(isWorksheetRegion(selfClosing, "cols")).toBe(false);
		pushWorksheetOpen(selfClosing, "sheetData", true);
		expect(isWorksheetRegion(selfClosing, "sheetData")).toBe(false);

		const wrongRoot = createWorksheetStructure();
		pushWorksheetOpen(wrongRoot, "wrapper", false);
		pushWorksheetOpen(wrongRoot, "worksheet", false);
		pushWorksheetOpen(wrongRoot, "sheetData", false);
		expect(wrongRoot.phase).toBe("closed");
		expect(isWorksheetRegion(wrongRoot, "sheetData")).toBe(false);
	});

	it("keeps containers nested under an extension outside canonical scope", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "extLst", false);
		pushWorksheetOpen(state, "ext", false);
		pushWorksheetOpen(state, "cols", false);
		expect(isWorksheetRegion(state, "cols")).toBe(false);
		pushWorksheetOpen(state, "sheetData", false);
		expect(isWorksheetRegion(state, "sheetData")).toBe(false);
		pushWorksheetClose(state, "sheetData");
		pushWorksheetClose(state, "cols");
		pushWorksheetClose(state, "ext");
		pushWorksheetClose(state, "extLst");

		// Once the extension's real direct-child close arrives, a following direct container works.
		pushWorksheetOpen(state, "sheetData", false);
		expect(isWorksheetRegion(state, "sheetData")).toBe(true);
	});

	it("does not let a nested same-name close release the direct owner", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "sheetData", false);
		pushWorksheetOpen(state, "sheetData", false);
		expect(pushWorksheetClose(state, "sheetData")).toBeUndefined();
		expect(isWorksheetRegion(state, "sheetData")).toBe(true);
		expect(pushWorksheetClose(state, "sheetData")).toBe("sheetData");
		expect(isWorksheetRegion(state, "sheetData")).toBe(false);
	});

	it("ignores unrelated closes and accepts sequential direct containers", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetClose(state, "unmatched");
		pushWorksheetOpen(state, "cols", false);
		pushWorksheetClose(state, "row");
		expect(isWorksheetRegion(state, "cols")).toBe(true);
		expect(pushWorksheetClose(state, "cols")).toBe("cols");
		pushWorksheetOpen(state, "cols", false);
		expect(isWorksheetRegion(state, "cols")).toBe(true);
	});

	it("admits only direct col, row, and cell children across extension barriers", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "cols", false);
		pushWorksheetOpen(state, "ext", false);
		expect(pushWorksheetOpen(state, "col", true)).toBeUndefined();
		pushWorksheetClose(state, "ext");
		expect(pushWorksheetOpen(state, "col", true)).toBe("col");
		pushWorksheetClose(state, "cols");

		pushWorksheetOpen(state, "sheetData", false);
		pushWorksheetOpen(state, "ext", false);
		expect(pushWorksheetOpen(state, "row", false)).toBeUndefined();
		pushWorksheetClose(state, "row");
		pushWorksheetClose(state, "ext");
		expect(pushWorksheetOpen(state, "row", false)).toBe("row");
		pushWorksheetOpen(state, "ext", false);
		expect(pushWorksheetOpen(state, "c", false)).toBeUndefined();
		expect(pushWorksheetClose(state, "c")).toBeUndefined();
		pushWorksheetClose(state, "ext");
		expect(pushWorksheetOpen(state, "c", true)).toBe("c");
	});

	it("lets canonical cell and row closes recover across an unclosed barrier", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "sheetData", false);
		expect(pushWorksheetOpen(state, "row", false)).toBe("row");
		expect(pushWorksheetOpen(state, "c", false)).toBe("c");
		pushWorksheetOpen(state, "ext", false);
		expect(pushWorksheetClose(state, "c")).toBe("c");
		expect(pushWorksheetOpen(state, "c", true)).toBe("c");
		pushWorksheetOpen(state, "ext", false);
		expect(pushWorksheetClose(state, "row")).toBe("row");
		expect(pushWorksheetOpen(state, "row", true)).toBe("row");
	});

	it("closes a truncated active region at the real worksheet root", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "sheetData", false);
		expect(pushWorksheetOpen(state, "row", false)).toBe("row");
		expect(pushWorksheetClose(state, "worksheet")).toBe("sheetData");
		expect(state.phase).toBe("closed");
		expect(pushWorksheetOpen(state, "row", true)).toBeUndefined();
	});

	it("does not confuse a nested worksheet close with the real root", () => {
		const state = createWorksheetStructure();
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetOpen(state, "extLst", false);
		pushWorksheetOpen(state, "worksheet", false);
		pushWorksheetClose(state, "worksheet");
		expect(state.phase).toBe("worksheet");
		pushWorksheetClose(state, "extLst");
		pushWorksheetOpen(state, "sheetData", false);
		expect(isWorksheetRegion(state, "sheetData")).toBe(true);
	});
});
