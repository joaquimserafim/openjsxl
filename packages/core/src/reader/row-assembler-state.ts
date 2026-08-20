import { type DecodeContext, decodeCell, decodeXstring, formatRef } from "../ooxml";
import type { Cell, Row } from "../types";
import { localName } from "../utils";
import type { XmlToken } from "../xml";
import { effectiveStyle, rowDefaultStyleFromToken, safeColumn } from "./worksheet-cell";
import {
	type ColumnStyleResolver,
	columnStyleFromToken,
	createColumnStyleResolver,
} from "./worksheet-column-styles";

// Internal row-assembler lifecycle and token transitions. They stay small enough to test directly
// without going through the XML tokenizer or exposing new package API.

export interface CellAssembly {
	open: boolean;
	ref: string;
	type: string | undefined;
	style: number | undefined;
	inline: boolean;
	value: string;
	inlineText: string;
	hasValue: boolean;
	inValue: boolean;
	inInline: boolean;
	textDepth: number;
	phoneticDepth: number;
}

export interface RowAssemblyState {
	readonly ctx: DecodeContext;
	inSheetData: boolean;
	inRow: boolean;
	lastRow: number;
	rowIndex: number;
	cells: Cell[];
	lastColumn: number;
	readonly columnStyles: ColumnStyleResolver;
	rowDefault: number | undefined;
	readonly cell: CellAssembly;
}

/** Create a closed, value-less cell state. Example: `{ open: false, value: "" }`. */
export function createCellAssembly(): CellAssembly {
	return {
		open: false,
		ref: "",
		type: undefined,
		style: undefined,
		inline: false,
		value: "",
		inlineText: "",
		hasValue: false,
		inValue: false,
		inInline: false,
		textDepth: 0,
		phoneticDepth: 0,
	};
}

/** Create the initial assembler state. Input: context; output: no open sheet, row, or cell. */
export function createRowAssemblyState(ctx: DecodeContext): RowAssemblyState {
	return {
		ctx,
		inSheetData: false,
		inRow: false,
		lastRow: 0,
		rowIndex: 0,
		cells: [],
		lastColumn: 0,
		columnStyles: createColumnStyleResolver(),
		rowDefault: undefined,
		cell: createCellAssembly(),
	};
}

/** Decode one completed inline run. Input `_x0041_`; output appends `A` and clears the buffer. */
export function flushInlineText(cell: CellAssembly): void {
	if (cell.inlineText === "") return;
	// Each inline <t> is one ST_Xstring context; escapes never join across rich-text runs.
	cell.value += decodeXstring(cell.inlineText);
	cell.inlineText = "";
}

/** Finish an open cell. Example: raw `B2 = "7"` appends the typed number cell `B2 = 7`. */
export function finishCell(state: RowAssemblyState): void {
	const { cell } = state;
	if (!cell.open) return;
	flushInlineText(cell);
	// A cached formula string's single <v> is also one ST_Xstring context.
	const value = cell.type === "str" ? decodeXstring(cell.value) : cell.value;
	state.cells.push(
		decodeCell(
			{
				ref: cell.ref,
				type: cell.type,
				value: cell.hasValue ? value : undefined,
				style: cell.style,
			},
			state.ctx,
		),
	);
	cell.open = false;
}

/** Finish the current row. Input: open row 2; output: row 2 and `inRow = false`. */
export function finishRow(state: RowAssemblyState): Row | undefined {
	finishCell(state);
	if (!state.inRow) return undefined;
	state.inRow = false;
	return { index: state.rowIndex, cells: state.cells };
}

/** Start a row, first recovering any open row. Input `<row r="2"/>`; output may contain two rows. */
export function startRow(
	state: RowAssemblyState,
	attrs: Readonly<Record<string, string>>,
	selfClosing: boolean,
): Row[] {
	const rows: Row[] = [];
	const previous = finishRow(state);
	if (previous !== undefined) rows.push(previous);

	const parsed = attrs.r !== undefined ? Number.parseInt(attrs.r, 10) : Number.NaN;
	state.rowIndex = Number.isInteger(parsed) && parsed > 0 ? parsed : state.lastRow + 1;
	state.lastRow = state.rowIndex;
	state.rowDefault = rowDefaultStyleFromToken(attrs);
	state.cells = [];
	state.lastColumn = 0;
	state.inRow = !selfClosing;
	if (selfClosing) rows.push({ index: state.rowIndex, cells: state.cells });
	return rows;
}

/** Start a cell. Example: row 5 after column C plus no `r` produces ref `D5`. */
export function startCell(
	state: RowAssemblyState,
	attrs: Readonly<Record<string, string>>,
	selfClosing: boolean,
): void {
	finishCell(state);

	let ref = attrs.r;
	let column: number | undefined;
	if (ref !== undefined) {
		column = safeColumn(ref);
		if (column !== undefined) state.lastColumn = column;
	} else {
		state.lastColumn += 1;
		column = state.lastColumn;
		ref = formatRef({ col: column, row: state.rowIndex });
	}

	const type = attrs.t;
	const { cell } = state;
	cell.open = true;
	cell.ref = ref;
	cell.type = type;
	cell.style = effectiveStyle(attrs.s, state.rowDefault, column, state.columnStyles);
	cell.inline = type === "inlineStr";
	cell.value = "";
	cell.inlineText = "";
	cell.hasValue = false;
	cell.inValue = false;
	cell.inInline = false;
	cell.textDepth = 0;
	cell.phoneticDepth = 0;
	if (selfClosing) finishCell(state);
}

/** Open the selected value channel. Example: inline `<is>` marks a present inline value. */
export function openValueChannel(cell: CellAssembly, name: string, selfClosing: boolean): void {
	if (cell.inline) {
		if (name === "is") {
			cell.hasValue = true;
			if (!selfClosing) cell.inInline = true;
		} else if (name === "t") {
			if (cell.inInline && !selfClosing) cell.textDepth++;
		} else if (name === "rPh" || name === "phoneticPr") {
			if (cell.inInline && !selfClosing) cell.phoneticDepth++;
		}
	} else if (name === "v") {
		cell.hasValue = true;
		if (!selfClosing) cell.inValue = true;
	}
}

/** Handle an opening token. Example: `<sheetData>` changes `inSheetData` to true and emits `[]`. */
export function pushOpen(
	state: RowAssemblyState,
	token: Extract<XmlToken, { kind: "open" }>,
): Row[] {
	const name = localName(token.name);
	if (name === "sheetData") {
		if (!token.selfClosing) state.inSheetData = true;
		return [];
	}
	// `<col>` precedes `<sheetData>`, so collect defaults before applying the sheet-data gate.
	if (name === "col") {
		const style = columnStyleFromToken(token.attrs);
		if (style !== undefined) state.columnStyles.add(style);
		return [];
	}
	if (!state.inSheetData) return [];

	if (name === "row") return startRow(state, token.attrs, token.selfClosing);
	if (!state.inRow) return [];
	if (name === "c") {
		startCell(state, token.attrs, token.selfClosing);
		return [];
	}
	if (state.cell.open) openValueChannel(state.cell, name, token.selfClosing);
	return [];
}

/** Append text only inside the active channel. Input numeric `<v>` + `"12"`; output value `"12"`. */
export function pushText(state: RowAssemblyState, value: string): void {
	const { cell } = state;
	if (!cell.open) return;
	const collect = cell.inline
		? cell.inInline && cell.textDepth > 0 && cell.phoneticDepth === 0
		: cell.inValue;
	if (!collect) return;
	if (cell.inline) cell.inlineText += value;
	else cell.value += value;
	cell.hasValue = true;
}

/** Handle a closing token. Example: `</row>` emits the completed row exactly once. */
export function pushClose(state: RowAssemblyState, nameWithPrefix: string): Row[] {
	const name = localName(nameWithPrefix);
	if (name === "sheetData") {
		const row = finishRow(state);
		state.inSheetData = false;
		return row === undefined ? [] : [row];
	}
	if (name === "row") {
		if (!state.inRow) return [];
		const row = finishRow(state);
		return row === undefined ? [] : [row];
	}

	const { cell } = state;
	if (!cell.open) return [];
	if (name === "c") finishCell(state);
	else if (name === "v") cell.inValue = false;
	else if (name === "is") cell.inInline = false;
	else if (name === "t") {
		if (cell.textDepth > 0) {
			cell.textDepth--;
			if (cell.textDepth === 0) flushInlineText(cell);
		}
	} else if (name === "rPh" || name === "phoneticPr") {
		if (cell.phoneticDepth > 0) cell.phoneticDepth--;
	}
	return [];
}

/** Flush end-of-input. Input: one open row; output: that row, then `[]` on the next flush. */
export function flushAssembler(state: RowAssemblyState): Row[] {
	const row = finishRow(state);
	return row === undefined ? [] : [row];
}

/** Dispatch one token. Input: `</row>`; output: the completed row, when one is open. */
export function pushToken(state: RowAssemblyState, token: XmlToken): Row[] {
	if (token.kind === "open") return pushOpen(state, token);
	if (token.kind === "text") {
		pushText(state, token.value);
		return [];
	}
	return pushClose(state, token.name);
}
