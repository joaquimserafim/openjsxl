import { localName } from "../utils";
import {
	clearDirectChild,
	isSheetRoot,
	pushColsOpen,
	pushRegionChildClose,
	pushSheetDataOpen,
	type WorksheetClose,
	type WorksheetOpen,
	type WorksheetRegion,
	type WorksheetStructure,
} from "./worksheet-structure-helpers";

export type {
	WorksheetClose,
	WorksheetOpen,
	WorksheetRegion,
	WorksheetStructure,
} from "./worksheet-structure-helpers";

/**
 * Create document scope before its root.
 * Example: input `()` returns `phase: "before-root"`, no root name, no owners/barriers, and zero
 * counters.
 */
export function createWorksheetStructure(): WorksheetStructure {
	return {
		phase: "before-root",
		rootName: undefined,
		rootSameNameDepth: 0,
		directChild: undefined,
		sameNameDepth: 0,
		region: undefined,
		regionChild: undefined,
		rowChild: undefined,
		cellBarrier: undefined,
		blockedRows: 0,
		blockedCells: 0,
	};
}

/**
 * Test whether one canonical direct container owns the token stream.
 * Examples: active `cols` + input `"cols"` returns `true`, including while a nested barrier is
 * open; input `"sheetData"`, a closed state, or a noncanonical direct-child owner returns `false`.
 */
export function isWorksheetRegion(state: WorksheetStructure, region: WorksheetRegion): boolean {
	return state.phase === "worksheet" && state.region === region;
}

/**
 * Test whether canonical sheetData currently owns an admitted row.
 * Examples: `sheetData` + `regionChild.kind === "row"` returns `true`, including while a lower
 * cell/extension barrier is open; no row, an extension owning sheetData's direct-child position,
 * `cols`, or a closed worksheet returns `false`.
 */
export function isWorksheetRow(state: WorksheetStructure): boolean {
	return state.region === "sheetData" && state.regionChild?.kind === "row";
}

/**
 * Track one opening element in constant space and return an admitted style-bearing element.
 * Examples:
 * - before-root + `<worksheet>`, `<x:worksheet>` or `<xm:macrosheet>` returns nothing, enters
 *   worksheet scope and records the root's local name (F11.32);
 * - before-root + `<worksheet/>`, `<xm:macrosheet/>` or another root returns nothing and permanently
 *   closes scope;
 * - any opening after closure returns nothing and changes no state;
 * - `<cols>`, `<sheetData>`, or another non-self-closing direct child returns nothing and owns the
 *   direct-child position; self-closing direct children return nothing and own nothing;
 * - below canonical `<cols>`, only a direct `<col>` returns `"col"`;
 * - below canonical `<sheetData>`, only an admitted direct `<row>` or row child `<c>` returns
 *   `"row"` or `"c"`;
 * - a non-self-closing fake `<row>`/`<c>` below a barrier returns nothing and increments its
 *   blocked counter; its self-closing form returns nothing without incrementing;
 * - `<sheetData><ext><row>` and every opening under a noncanonical direct child return nothing;
 * - a nested non-self-closing root of the recorded name raises the root-name depth, while another
 *   opening with the current direct-child name raises the direct-child depth; both return nothing
 *   and prevent the corresponding inner close from releasing the outer owner.
 */
export function pushWorksheetOpen(
	state: WorksheetStructure,
	nameWithPrefix: string,
	selfClosing: boolean,
): WorksheetOpen | undefined {
	const name = localName(nameWithPrefix);
	if (state.phase === "before-root") {
		if (isSheetRoot(name) && !selfClosing) {
			state.phase = "worksheet";
			state.rootName = name;
		} else {
			state.phase = "closed";
		}
		return undefined;
	}
	if (state.phase !== "worksheet") return undefined;
	if (name === state.rootName && !selfClosing) state.rootSameNameDepth++;

	if (state.directChild === undefined) {
		if (selfClosing) return undefined;
		state.directChild = name;
		state.sameNameDepth = 0;
		state.region = name === "cols" || name === "sheetData" ? name : undefined;
		state.regionChild = undefined;
		state.rowChild = undefined;
		state.cellBarrier = undefined;
		state.blockedRows = 0;
		state.blockedCells = 0;
		return undefined;
	}

	// Same-name nesting prevents an inner close from releasing the worksheet's direct owner.
	if (!selfClosing && name === state.directChild) state.sameNameDepth++;
	if (state.region === "cols") return pushColsOpen(state, name, selfClosing);
	if (state.region === "sheetData") return pushSheetDataOpen(state, name, selfClosing);
	return undefined;
}

/**
 * Track one closing element and return the canonical boundary that actually ended, if any.
 * Examples:
 * - before-root/closed state or an unrelated close returns nothing and changes no owner;
 * - prefixed names such as `</x:row>` have the same result as their local name;
 * - `</c>` / `</row>` for admitted owners returns `"c"` / `"row"` and releases that scope;
 * - an extension/barrier close returns nothing, clearing it only at its matching outer close;
 * - a well-formed blocked fake `</row>` / `</c>` returns nothing and decrements its blocked
 *   counter, so it cannot close the admitted owner;
 * - the inner close of a same-name direct child returns nothing and decrements its depth;
 * - the outer `</cols>` / `</sheetData>` returns `"cols"` / `"sheetData"` and clears the direct
 *   child; a noncanonical direct-child close returns nothing and clears its barrier;
 * - a nested close of the recorded root name (`</worksheet>` or `</xm:macrosheet>`) returns nothing
 *   and decrements `rootSameNameDepth`;
 * - the real root close returns the active `"cols"` or `"sheetData"` boundary, or nothing when no
 *   canonical region is active, then closes the state permanently; a close of the other sheet-root
 *   name is an ordinary close.
 */
export function pushWorksheetClose(
	state: WorksheetStructure,
	nameWithPrefix: string,
): WorksheetClose | undefined {
	if (state.phase !== "worksheet") return undefined;
	const name = localName(nameWithPrefix);

	if (name === state.rootName && state.rootSameNameDepth === 0) {
		const ended = state.region;
		clearDirectChild(state);
		state.phase = "closed";
		return ended;
	}

	const boundary = pushRegionChildClose(state, name);
	if (state.directChild !== undefined && name === state.directChild) {
		if (state.sameNameDepth > 0) state.sameNameDepth--;
		else {
			const ended = state.region;
			clearDirectChild(state);
			if (ended !== undefined) return ended;
		}
	}

	if (name === state.rootName && state.rootSameNameDepth > 0) state.rootSameNameDepth--;
	return boundary;
}
