import { describe, expect, it } from "vitest";
import { createWorksheetStructure } from "../worksheet-structure";
import {
	clearDirectChild,
	closesOwner,
	createRegionChild,
	isCellContent,
	pushColsOpen,
	pushNestedOpen,
	pushRegionChildClose,
	pushRowOpen,
	pushSheetDataOpen,
	type RegionChild,
	type WorksheetStructure,
} from "../worksheet-structure-helpers";

// These tests exercise every extracted state transition directly. Literal owners isolate each
// helper from createRegionChild(), while createWorksheetStructure supplies the neutral state.

describe("worksheet structure helpers — owner primitives", () => {
	it("creates every owner kind at depth zero with separate identity", () => {
		expect(createRegionChild("col", "col")).toEqual({
			name: "col",
			kind: "col",
			sameNameDepth: 0,
		});
		expect(createRegionChild("row", "row")).toEqual({
			name: "row",
			kind: "row",
			sameNameDepth: 0,
		});
		expect(createRegionChild("c", "c")).toEqual({ name: "c", kind: "c", sameNameDepth: 0 });
		expect(createRegionChild("ext", "other")).toEqual({
			name: "ext",
			kind: "other",
			sameNameDepth: 0,
		});
		expect(createRegionChild("row", "row")).not.toBe(createRegionChild("row", "row"));
	});

	it("tracks only matching non-self-closing nested owners", () => {
		const owner: RegionChild = { name: "ext", kind: "other", sameNameDepth: 0 };
		pushNestedOpen(owner, "row", false);
		expect(owner).toEqual({ name: "ext", kind: "other", sameNameDepth: 0 });
		pushNestedOpen(owner, "ext", true);
		expect(owner.sameNameDepth).toBe(0);
		pushNestedOpen(owner, "ext", false);
		expect(owner).toEqual({ name: "ext", kind: "other", sameNameDepth: 1 });
	});

	it("rejects unrelated closes, consumes nested closes, then releases the owner", () => {
		const owner: RegionChild = { name: "ext", kind: "other", sameNameDepth: 2 };
		expect(closesOwner(owner, "row")).toBe(false);
		expect(owner.sameNameDepth).toBe(2);
		expect(closesOwner(owner, "ext")).toBe(false);
		expect(owner.sameNameDepth).toBe(1);
		expect(closesOwner(owner, "ext")).toBe(false);
		expect(owner.sameNameDepth).toBe(0);
		expect(closesOwner(owner, "ext")).toBe(true);
		expect(owner).toEqual({ name: "ext", kind: "other", sameNameDepth: 0 });
	});

	it("recognizes only the allowed cell-content local names", () => {
		for (const name of ["f", "v", "is", "t", "r", "rPr", "rPh", "phoneticPr"])
			expect(isCellContent(name)).toBe(true);
		for (const name of ["ext", "extLst", "c", "row", "", "V"])
			expect(isCellContent(name)).toBe(false);
	});
});

describe("worksheet structure helpers — opening transitions", () => {
	it("admits direct cols children and distinguishes retained owners", () => {
		const selfClosingCol = createWorksheetStructure();
		expect(pushColsOpen(selfClosingCol, "col", true)).toBe("col");
		expect(selfClosingCol.regionChild).toBeUndefined();

		const openCol = createWorksheetStructure();
		expect(pushColsOpen(openCol, "col", false)).toBe("col");
		expect(openCol.regionChild).toEqual({ name: "col", kind: "col", sameNameDepth: 0 });

		const selfClosingExtension = createWorksheetStructure();
		expect(pushColsOpen(selfClosingExtension, "ext", true)).toBeUndefined();
		expect(selfClosingExtension.regionChild).toBeUndefined();

		const openExtension = createWorksheetStructure();
		expect(pushColsOpen(openExtension, "ext", false)).toBeUndefined();
		expect(openExtension.regionChild).toEqual({
			name: "ext",
			kind: "other",
			sameNameDepth: 0,
		});
	});

	it("blocks every cols opening below an existing owner and tracks only matching depth", () => {
		const state = createWorksheetStructure();
		state.regionChild = { name: "ext", kind: "other", sameNameDepth: 0 };
		expect(pushColsOpen(state, "col", false)).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(0);
		expect(pushColsOpen(state, "ext", true)).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(0);
		expect(pushColsOpen(state, "ext", false)).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(1);
	});

	it("admits direct row cells and retains only non-self-closing owners", () => {
		const selfClosingCell = createWorksheetStructure();
		expect(pushRowOpen(selfClosingCell, "c", true)).toBe("c");
		expect(selfClosingCell.rowChild).toBeUndefined();

		const openCell = createWorksheetStructure();
		expect(pushRowOpen(openCell, "c", false)).toBe("c");
		expect(openCell.rowChild).toEqual({ name: "c", kind: "c", sameNameDepth: 0 });

		const selfClosingExtension = createWorksheetStructure();
		expect(pushRowOpen(selfClosingExtension, "ext", true)).toBeUndefined();
		expect(selfClosingExtension.rowChild).toBeUndefined();

		const openExtension = createWorksheetStructure();
		expect(pushRowOpen(openExtension, "ext", false)).toBeUndefined();
		expect(openExtension.rowChild).toEqual({
			name: "ext",
			kind: "other",
			sameNameDepth: 0,
		});
	});

	it("recovers a missing cell close and follows the replacement cell form", () => {
		const openReplacement = createWorksheetStructure();
		const prior: RegionChild = { name: "c", kind: "c", sameNameDepth: 0 };
		openReplacement.rowChild = prior;
		expect(pushRowOpen(openReplacement, "c", false)).toBe("c");
		expect(openReplacement.rowChild).toEqual({ name: "c", kind: "c", sameNameDepth: 0 });
		expect(openReplacement.rowChild).not.toBe(prior);

		const selfClosingReplacement = createWorksheetStructure();
		selfClosingReplacement.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		expect(pushRowOpen(selfClosingReplacement, "c", true)).toBe("c");
		expect(selfClosingReplacement.rowChild).toBeUndefined();
	});

	it("counts fake cells under cell and row-child barriers", () => {
		const cellBarrier = createWorksheetStructure();
		cellBarrier.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		cellBarrier.cellBarrier = { name: "ext", kind: "other", sameNameDepth: 0 };
		expect(pushRowOpen(cellBarrier, "c", false)).toBeUndefined();
		expect(cellBarrier.blockedCells).toBe(1);
		expect(pushRowOpen(cellBarrier, "c", true)).toBeUndefined();
		expect(cellBarrier.blockedCells).toBe(1);
		expect(pushRowOpen(cellBarrier, "ext", false)).toBeUndefined();
		expect(cellBarrier.cellBarrier.sameNameDepth).toBe(1);

		const rowBarrier = createWorksheetStructure();
		rowBarrier.rowChild = { name: "ext", kind: "other", sameNameDepth: 0 };
		expect(pushRowOpen(rowBarrier, "c", false)).toBeUndefined();
		expect(rowBarrier.blockedCells).toBe(1);
		expect(pushRowOpen(rowBarrier, "c", true)).toBeUndefined();
		expect(rowBarrier.blockedCells).toBe(1);
		expect(pushRowOpen(rowBarrier, "ext", false)).toBeUndefined();
		expect(rowBarrier.rowChild.sameNameDepth).toBe(1);
	});

	it("allows known cell content and barriers only unknown non-self-closing content", () => {
		const known = createWorksheetStructure();
		known.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		for (const name of ["f", "v", "is", "t", "r", "rPr", "rPh", "phoneticPr"])
			expect(pushRowOpen(known, name, false)).toBeUndefined();
		expect(known.cellBarrier).toBeUndefined();
		expect(pushRowOpen(known, "ext", true)).toBeUndefined();
		expect(known.cellBarrier).toBeUndefined();
		expect(pushRowOpen(known, "ext", false)).toBeUndefined();
		expect(known.cellBarrier).toEqual({ name: "ext", kind: "other", sameNameDepth: 0 });
	});

	it("admits direct sheetData rows and retains only non-self-closing owners", () => {
		const selfClosingRow = createWorksheetStructure();
		expect(pushSheetDataOpen(selfClosingRow, "row", true)).toBe("row");
		expect(selfClosingRow.regionChild).toBeUndefined();

		const openRow = createWorksheetStructure();
		expect(pushSheetDataOpen(openRow, "row", false)).toBe("row");
		expect(openRow.regionChild).toEqual({ name: "row", kind: "row", sameNameDepth: 0 });

		const selfClosingExtension = createWorksheetStructure();
		expect(pushSheetDataOpen(selfClosingExtension, "ext", true)).toBeUndefined();
		expect(selfClosingExtension.regionChild).toBeUndefined();

		const openExtension = createWorksheetStructure();
		expect(pushSheetDataOpen(openExtension, "ext", false)).toBeUndefined();
		expect(openExtension.regionChild).toEqual({
			name: "ext",
			kind: "other",
			sameNameDepth: 0,
		});
	});

	it("recovers a missing row close and resets lower state and counters", () => {
		const openReplacement = createWorksheetStructure();
		openReplacement.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		openReplacement.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		openReplacement.blockedRows = 3;
		openReplacement.blockedCells = 4;
		expect(pushSheetDataOpen(openReplacement, "row", false)).toBe("row");
		expect(openReplacement.regionChild).toEqual({
			name: "row",
			kind: "row",
			sameNameDepth: 0,
		});
		expect(openReplacement.rowChild).toBeUndefined();
		expect(openReplacement.cellBarrier).toBeUndefined();
		expect(openReplacement.blockedRows).toBe(0);
		expect(openReplacement.blockedCells).toBe(0);

		const selfClosingReplacement = createWorksheetStructure();
		selfClosingReplacement.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		selfClosingReplacement.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		expect(pushSheetDataOpen(selfClosingReplacement, "row", true)).toBe("row");
		expect(selfClosingReplacement.regionChild).toBeUndefined();
		expect(selfClosingReplacement.rowChild).toBeUndefined();
	});

	it("blocks nested rows and prefers the row-child barrier", () => {
		const rowBarrier = createWorksheetStructure();
		rowBarrier.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		rowBarrier.rowChild = { name: "row", kind: "other", sameNameDepth: 0 };
		rowBarrier.cellBarrier = { name: "row", kind: "other", sameNameDepth: 0 };
		expect(pushSheetDataOpen(rowBarrier, "row", false)).toBeUndefined();
		expect(rowBarrier.blockedRows).toBe(1);
		expect(rowBarrier.rowChild.sameNameDepth).toBe(1);
		expect(rowBarrier.cellBarrier.sameNameDepth).toBe(0);
		expect(pushSheetDataOpen(rowBarrier, "row", true)).toBeUndefined();
		expect(rowBarrier.blockedRows).toBe(1);
		expect(rowBarrier.rowChild.sameNameDepth).toBe(1);
		expect(rowBarrier.cellBarrier.sameNameDepth).toBe(0);

		const cellBarrier = createWorksheetStructure();
		cellBarrier.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		cellBarrier.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		cellBarrier.cellBarrier = { name: "row", kind: "other", sameNameDepth: 0 };
		expect(pushSheetDataOpen(cellBarrier, "row", false)).toBeUndefined();
		expect(cellBarrier.blockedRows).toBe(1);
		expect(cellBarrier.cellBarrier.sameNameDepth).toBe(1);
	});

	it("delegates non-row openings from an active row", () => {
		const state = createWorksheetStructure();
		state.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		expect(pushSheetDataOpen(state, "c", false)).toBe("c");
		expect(state.rowChild).toEqual({ name: "c", kind: "c", sameNameDepth: 0 });
	});

	it("blocks openings under a sheetData child barrier", () => {
		const state = createWorksheetStructure();
		state.regionChild = { name: "ext", kind: "other", sameNameDepth: 0 };
		expect(pushSheetDataOpen(state, "row", false)).toBeUndefined();
		expect(state.blockedRows).toBe(1);
		expect(pushSheetDataOpen(state, "row", true)).toBeUndefined();
		expect(state.blockedRows).toBe(1);
		expect(pushSheetDataOpen(state, "ext", false)).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(1);
	});
});

describe("worksheet structure helpers — closing transitions", () => {
	it("does nothing when no region child owns the close", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		state.blockedRows = 2;
		expect(pushRegionChildClose(state, "row")).toBeUndefined();
		expect(state.blockedRows).toBe(2);
	});

	it("retains cols owners for unrelated and nested closes, then clears the outer owner", () => {
		const state = createWorksheetStructure();
		state.region = "cols";
		state.regionChild = { name: "ext", kind: "other", sameNameDepth: 1 };
		state.rootSameNameDepth = 2;
		expect(pushRegionChildClose(state, "row")).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(1);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(0);
		expect(state.rootSameNameDepth).toBe(2);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.regionChild).toBeUndefined();
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("consumes a blocked row close before closing the canonical row", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		state.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		state.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		state.cellBarrier = { name: "ext", kind: "other", sameNameDepth: 0 };
		state.blockedRows = 1;
		state.blockedCells = 2;
		state.rootSameNameDepth = 3;
		expect(pushRegionChildClose(state, "row")).toBeUndefined();
		expect(state.blockedRows).toBe(0);
		expect(state.regionChild?.kind).toBe("row");
		expect(state.rowChild?.kind).toBe("c");
		expect(state.cellBarrier?.name).toBe("ext");
		expect(state.blockedCells).toBe(2);
		expect(state.rootSameNameDepth).toBe(3);
		expect(pushRegionChildClose(state, "row")).toBe("row");
		expect(state.regionChild).toBeUndefined();
		expect(state.rowChild).toBeUndefined();
		expect(state.cellBarrier).toBeUndefined();
		expect(state.blockedRows).toBe(0);
		expect(state.blockedCells).toBe(0);
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("consumes a blocked cell close before closing the canonical cell", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		const rowOwner: RegionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		state.regionChild = rowOwner;
		state.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		state.cellBarrier = { name: "ext", kind: "other", sameNameDepth: 0 };
		state.blockedRows = 2;
		state.blockedCells = 1;
		state.rootSameNameDepth = 3;
		expect(pushRegionChildClose(state, "c")).toBeUndefined();
		expect(state.blockedCells).toBe(0);
		expect(state.regionChild).toBe(rowOwner);
		expect(state.rowChild?.kind).toBe("c");
		expect(state.cellBarrier?.name).toBe("ext");
		expect(state.blockedRows).toBe(2);
		expect(state.rootSameNameDepth).toBe(3);
		expect(pushRegionChildClose(state, "c")).toBe("c");
		expect(state.regionChild).toBe(rowOwner);
		expect(state.rowChild).toBeUndefined();
		expect(state.cellBarrier).toBeUndefined();
		expect(state.blockedRows).toBe(0);
		expect(state.blockedCells).toBe(0);
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("clears a sheetData child barrier only at its outer close", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		state.regionChild = { name: "ext", kind: "other", sameNameDepth: 1 };
		state.blockedRows = 2;
		state.blockedCells = 3;
		state.rootSameNameDepth = 4;
		expect(pushRegionChildClose(state, "other")).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(1);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.regionChild.sameNameDepth).toBe(0);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.regionChild).toBeUndefined();
		expect(state.blockedRows).toBe(0);
		expect(state.blockedCells).toBe(0);
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("clears a row-child barrier only at its outer close", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		state.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		state.rowChild = { name: "ext", kind: "other", sameNameDepth: 1 };
		state.blockedRows = 2;
		state.blockedCells = 3;
		state.rootSameNameDepth = 4;
		expect(pushRegionChildClose(state, "other")).toBeUndefined();
		expect(state.rowChild.sameNameDepth).toBe(1);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.rowChild.sameNameDepth).toBe(0);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.rowChild).toBeUndefined();
		expect(state.regionChild?.kind).toBe("row");
		expect(state.blockedRows).toBe(0);
		expect(state.blockedCells).toBe(0);
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("clears a cell barrier only at its outer close", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		state.regionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		state.rowChild = { name: "c", kind: "c", sameNameDepth: 0 };
		state.cellBarrier = { name: "ext", kind: "other", sameNameDepth: 1 };
		state.blockedRows = 2;
		state.blockedCells = 3;
		state.rootSameNameDepth = 4;
		expect(pushRegionChildClose(state, "other")).toBeUndefined();
		expect(state.cellBarrier.sameNameDepth).toBe(1);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.cellBarrier.sameNameDepth).toBe(0);
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.cellBarrier).toBeUndefined();
		expect(state.regionChild?.kind).toBe("row");
		expect(state.rowChild?.kind).toBe("c");
		expect(state.blockedRows).toBe(0);
		expect(state.blockedCells).toBe(0);
		expect(state.rootSameNameDepth).toBe(0);
	});

	it("keeps a canonical row unchanged for an unrelated close", () => {
		const state = createWorksheetStructure();
		state.region = "sheetData";
		const rowOwner: RegionChild = { name: "row", kind: "row", sameNameDepth: 0 };
		state.regionChild = rowOwner;
		expect(pushRegionChildClose(state, "ext")).toBeUndefined();
		expect(state.regionChild).toBe(rowOwner);
	});
});

describe("worksheet structure helpers — direct-child reset", () => {
	it("clears every nested field and counter while preserving each phase", () => {
		const phases: WorksheetStructure["phase"][] = ["before-root", "worksheet", "closed"];
		for (const phase of phases) {
			const state = createWorksheetStructure();
			state.phase = phase;
			state.rootSameNameDepth = 8;
			state.directChild = "sheetData";
			state.sameNameDepth = 7;
			state.region = "sheetData";
			state.regionChild = { name: "row", kind: "row", sameNameDepth: 6 };
			state.rowChild = { name: "c", kind: "c", sameNameDepth: 5 };
			state.cellBarrier = { name: "ext", kind: "other", sameNameDepth: 4 };
			state.blockedRows = 3;
			state.blockedCells = 2;

			clearDirectChild(state);

			expect(state).toEqual({
				phase,
				rootSameNameDepth: 0,
				directChild: undefined,
				sameNameDepth: 0,
				region: undefined,
				regionChild: undefined,
				rowChild: undefined,
				cellBarrier: undefined,
				blockedRows: 0,
				blockedCells: 0,
			});
		}
	});
});
