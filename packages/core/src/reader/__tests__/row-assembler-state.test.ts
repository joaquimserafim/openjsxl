import { describe, expect, it } from "vitest";
import type { XmlToken } from "../../xml";
import {
	createCellAssembly,
	createRowAssemblyState,
	finishCell,
	finishRow,
	flushAssembler,
	flushInlineText,
	openValueChannel,
	pushClose,
	pushOpen,
	pushText,
	pushToken,
	startCell,
	startRow,
} from "../row-assembler-state";

// Tokens are built directly so these tests cover state transitions, not XML tokenization.
const open = (
	name: string,
	attrs: Readonly<Record<string, string>> = {},
	selfClosing = false,
): Extract<XmlToken, { kind: "open" }> => ({ kind: "open", name, attrs, selfClosing });
const text = (value: string): XmlToken => ({ kind: "text", value });
const close = (name: string): XmlToken => ({ kind: "close", name });

describe("row assembler state — lifecycle", () => {
	it("creates independent closed states", () => {
		const left = createRowAssemblyState({ sharedStrings: [] });
		const right = createRowAssemblyState({ sharedStrings: [] });
		expect(left).toMatchObject({
			inSheetData: false,
			inRow: false,
			lastRow: 0,
			rowIndex: 0,
			lastColumn: 0,
			rowDefault: undefined,
		});
		expect(left.cells).not.toBe(right.cells);
		expect(left.columnStyles).not.toBe(right.columnStyles);
		expect(left.cell).not.toBe(right.cell);
		expect(createCellAssembly()).toMatchObject({ open: false, value: "", hasValue: false });
	});

	it("decodes one inline Xstring run, clears it, and is idempotent", () => {
		const cell = createCellAssembly();
		cell.value = "prefix:";
		cell.inlineText = "_x0041_";
		flushInlineText(cell);
		expect(cell.value).toBe("prefix:A");
		expect(cell.inlineText).toBe("");
		flushInlineText(cell);
		expect(cell.value).toBe("prefix:A");
	});

	it("finishes a cached-string cell exactly once", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		Object.assign(state.cell, {
			open: true,
			ref: "B2",
			type: "str",
			value: "_x0041_",
			hasValue: true,
		});
		finishCell(state);
		expect(state.cells).toEqual([{ ref: "B2", type: "string", value: "A" }]);
		expect(state.cell.open).toBe(false);
		finishCell(state);
		expect(state.cells).toHaveLength(1);
	});

	it("finishes a row once and keeps emitted rows stable", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.inRow = true;
		state.rowIndex = 2;
		state.lastRow = 2;
		state.cells.push({ ref: "A2", type: "number", value: 2 });
		const emitted = finishRow(state);
		expect(emitted).toEqual({
			index: 2,
			cells: [{ ref: "A2", type: "number", value: 2 }],
		});
		expect(finishRow(state)).toBeUndefined();

		startRow(state, { r: "3" }, false);
		startCell(state, { r: "A3" }, true);
		expect(emitted?.cells).toEqual([{ ref: "A2", type: "number", value: 2 }]);
	});

	it("recovers an open row before a new self-closing row and resets row state", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.inRow = true;
		state.rowIndex = 1;
		state.lastRow = 1;
		state.cells.push({ ref: "A1", type: "empty", value: null });
		state.lastColumn = 8;
		state.rowDefault = 9;

		expect(startRow(state, { r: "2", s: "5", customFormat: "1" }, true)).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "empty", value: null }] },
			{ index: 2, cells: [] },
		]);
		expect(state).toMatchObject({
			inRow: false,
			rowIndex: 2,
			lastRow: 2,
			lastColumn: 0,
			rowDefault: 5,
			cells: [],
		});
	});

	it("starts positional cells, preserves explicit refs, and resets value state", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.rowIndex = 5;
		state.lastColumn = 3;
		Object.assign(state.cell, {
			value: "stale",
			inlineText: "stale",
			hasValue: true,
			inValue: true,
			inInline: true,
			textDepth: 2,
			phoneticDepth: 1,
		});

		startCell(state, {}, false);
		expect(state.cell).toMatchObject({
			open: true,
			ref: "D5",
			value: "",
			inlineText: "",
			hasValue: false,
			inValue: false,
			inInline: false,
			textDepth: 0,
			phoneticDepth: 0,
		});
		startCell(state, { r: "F99" }, true); // first finishes D5, then emits F99
		expect(state.lastColumn).toBe(6);
		startCell(state, {}, false);
		expect(state.cell.ref).toBe("G5");
		startCell(state, { r: "bad" }, true); // malformed explicit refs do not advance the column
		startCell(state, {}, false);
		expect(state.cell.ref).toBe("H5");
		expect(state.cells.map((cell) => cell.ref)).toEqual(["D5", "F99", "G5", "bad"]);
	});
});

describe("row assembler state — token transitions", () => {
	it("opens only the selected value channel and collects eligible text", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.cell.open = true;
		openValueChannel(state.cell, "is", false);
		pushText(state, "ignored");
		expect(state.cell.value).toBe("");
		openValueChannel(state.cell, "v", false);
		pushText(state, "12");
		expect(state.cell).toMatchObject({ hasValue: true, inValue: true, value: "12" });

		Object.assign(state.cell, createCellAssembly(), { open: true, inline: true });
		openValueChannel(state.cell, "is", false);
		openValueChannel(state.cell, "t", false);
		pushText(state, "漢字");
		openValueChannel(state.cell, "rPh", false);
		pushText(state, "かんじ");
		expect(state.cell.inlineText).toBe("漢字");
		pushClose(state, "rPh");
		pushClose(state, "t");
		expect(state.cell).toMatchObject({ value: "漢字", inlineText: "", textDepth: 0 });
	});

	it("gates rows on sheetData and captures prefixed column styles first", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		expect(pushOpen(state, open("row", { r: "9" }, true))).toEqual([]);
		pushOpen(state, open("x:col", { min: "2", max: "3", style: "7" }, true));
		// Input columns 2..3/style 7; output is a direct lookup with no retained range scan.
		expect(state.columnStyles.styleFor(2)).toBe(7);
		expect(state.columnStyles.styleFor(4)).toBeUndefined();
		pushOpen(state, open("x:sheetData", {}, true));
		expect(state.inSheetData).toBe(false);
		pushOpen(state, open("x:sheetData"));
		expect(state.inSheetData).toBe(true);
		expect(pushOpen(state, open("x:row", { r: "2" }, true))).toEqual([{ index: 2, cells: [] }]);
	});

	it("closes rows and sheetData at the exact token", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.inSheetData = true;
		state.inRow = true;
		state.rowIndex = 3;
		startCell(state, { r: "A3" }, true);
		expect(pushClose(state, "x:row")).toEqual([
			{ index: 3, cells: [{ ref: "A3", type: "empty", value: null }] },
		]);
		expect(pushClose(state, "row")).toEqual([]);

		state.inRow = true;
		state.rowIndex = 4;
		state.cells = [];
		expect(pushClose(state, "sheetData")).toEqual([{ index: 4, cells: [] }]);
		expect(state.inSheetData).toBe(false);
	});

	it("flushes an open row once at end-of-input", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		state.inRow = true;
		state.rowIndex = 6;
		startCell(state, { r: "A6" }, true);
		expect(flushAssembler(state)).toEqual([
			{ index: 6, cells: [{ ref: "A6", type: "empty", value: null }] },
		]);
		expect(flushAssembler(state)).toEqual([]);
	});

	it("dispatches open, text, and close tokens", () => {
		const state = createRowAssemblyState({ sharedStrings: [] });
		for (const token of [
			open("sheetData"),
			open("row", { r: "1" }),
			open("c", { r: "A1" }),
			open("v"),
			text("7"),
			close("v"),
			close("c"),
		]) {
			expect(pushToken(state, token)).toEqual([]);
		}
		expect(pushToken(state, close("row"))).toEqual([
			{ index: 1, cells: [{ ref: "A1", type: "number", value: 7 }] },
		]);
	});
});
