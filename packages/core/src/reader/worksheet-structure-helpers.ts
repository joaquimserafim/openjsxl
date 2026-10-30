export type WorksheetRegion = "cols" | "sheetData";
/**
 * The direct children of an admitted cell that carry its content: the formula, the cell value and the
 * inline string (F11.26, F11.38). Only these, outside any extension barrier, open a channel.
 */
export type CellChild = "f" | "v" | "is";
export type WorksheetOpen = "col" | "row" | "c" | CellChild;
export type WorksheetClose = WorksheetRegion | "row" | "c" | CellChild;

export interface RegionChild {
	name: string;
	sameNameDepth: number;
	kind: WorksheetOpen | "other";
}

export interface WorksheetStructure {
	phase: "before-root" | "worksheet" | "closed";
	/**
	 * Local name of the document root the gate admitted (`worksheet`, or `macrosheet` for an Excel
	 * 4.0 macro sheet, F11.32); nested same-name roots and the root's close follow it.
	 */
	rootName: SheetRootName | undefined;
	rootSameNameDepth: number;
	directChild: string | undefined;
	sameNameDepth: number;
	region: WorksheetRegion | undefined;
	regionChild: RegionChild | undefined;
	rowChild: RegionChild | undefined;
	cellBarrier: RegionChild | undefined;
	/**
	 * The admitted cell's open direct child, when it is cell content (F11.26): `f`, `v` and `is` carry
	 * the cell's formula and value; any other known content name is tracked as `"other"`.
	 */
	cellChild: RegionChild | undefined;
	blockedRows: number;
	blockedCells: number;
}

/** The local names of the document roots whose content the gate reads (F11.32). */
export type SheetRootName = "worksheet" | "macrosheet";

/**
 * Whether a document root's local name (its prefix already removed) holds worksheet content the gate
 * reads: `worksheet`, or `macrosheet`, the root of both Excel 4.0 macro-sheet parts (MS-OFFMACRO2
 * 2.2.1.2; F11.32). Chart and dialog sheets, and any other root, do not.
 *
 * @example isSheetRoot("worksheet") // true
 * @example isSheetRoot("macrosheet") // true (xlMacrosheet and xlIntlMacrosheet parts)
 * @example isSheetRoot("chartsheet") // false
 * @example isSheetRoot("xm:macrosheet") // false (pass the local name)
 */
export function isSheetRoot(name: string): name is SheetRootName {
	return name === "worksheet" || name === "macrosheet";
}

/**
 * Create one fixed-size owner record.
 * Example: input `("row", "row")` returns `{ name: "row", kind: "row", sameNameDepth: 0 }`;
 * `"col"`, `"c"`, `"f"`, `"v"`, `"is"` and `"other"` kinds produce the same shape with that kind.
 */
export function createRegionChild(name: string, kind: RegionChild["kind"]): RegionChild {
	return { name, sameNameDepth: 0, kind };
}

/**
 * Record an opening element nested beneath an owner.
 * Examples: owner `ext`, input `("ext", false)` increments `sameNameDepth`; `("row", false)` and
 * `("ext", true)` leave it unchanged. The function returns no value.
 */
export function pushNestedOpen(owner: RegionChild, name: string, selfClosing: boolean): void {
	if (!selfClosing && name === owner.name) owner.sameNameDepth++;
}

/**
 * Decide whether a close releases its recorded owner.
 * Examples: a different name returns `false`; a matching name at depth 2 returns `false` and
 * decrements to 1; a matching name at depth 0 returns `true` without mutating the owner.
 */
export function closesOwner(owner: RegionChild, name: string): boolean {
	if (name !== owner.name) return false;
	if (owner.sameNameDepth > 0) {
		owner.sameNameDepth--;
		return false;
	}
	return true;
}

/**
 * Recognize elements allowed inside an active cell without opening an extension barrier.
 * Examples: `"f"`, `"v"`, `"is"`, `"t"`, `"r"`, `"rPr"`, `"rPh"`, and `"phoneticPr"`
 * return `true`; `"ext"`, `"extLst"`, `"c"`, and unknown names return `false`.
 */
export function isCellContent(name: string): boolean {
	return (
		name === "f" ||
		name === "v" ||
		name === "is" ||
		name === "t" ||
		name === "r" ||
		name === "rPr" ||
		name === "rPh" ||
		name === "phoneticPr"
	);
}

/**
 * Process one child opening inside canonical `<cols>`.
 * Examples:
 * - no owner + `<col/>` returns `"col"` and retains no owner;
 * - no owner + `<col>` returns `"col"` and owns that col until `</col>`;
 * - no owner + `<ext>` returns nothing and opens an extension barrier;
 * - no owner + `<ext/>` returns nothing and opens no barrier;
 * - any opening below an existing col/barrier returns nothing; a matching non-self-closing name
 *   only increases that owner's same-name depth.
 */
export function pushColsOpen(
	state: WorksheetStructure,
	name: string,
	selfClosing: boolean,
): WorksheetOpen | undefined {
	const owner = state.regionChild;
	if (owner !== undefined) {
		pushNestedOpen(owner, name, selfClosing);
		return undefined;
	}
	if (name === "col") {
		if (!selfClosing) state.regionChild = createRegionChild(name, "col");
		return "col";
	}
	if (!selfClosing) state.regionChild = createRegionChild(name, "other");
	return undefined;
}

/**
 * Process one opening beneath an admitted canonical row.
 * Examples:
 * - no child + `<c/>` returns `"c"`; `<c>` also records the active cell;
 * - no child + `<ext>` returns nothing and opens a row-child barrier; `<ext/>` opens none;
 * - active cell, no barrier and no open direct child + `<f>`, `<v>` or `<is>` returns `"f"`, `"v"`
 *   or `"is"` and, when not self-closing, records it as the cell's direct child (F11.26);
 * - the same with other known content (`<t>`, `<r>`, `<rPh>`, …) returns nothing and records an
 *   `"other"` direct child;
 * - known content inside the open direct child returns nothing and only tracks same-name nesting,
 *   so a nested `<v>` or `<f>` is not the cell's own;
 * - active cell + unknown `<ext>` returns nothing and opens a cell barrier, at any depth;
 * - active cell + unknown `<ext/>` returns nothing and opens no barrier;
 * - active cell + another `<c>` returns `"c"` and recovers the missing prior `</c>`; the new
 *   non-self-closing cell replaces the owner, while `<c/>` leaves no active cell;
 * - another `<c>` inside a barrier returns nothing and increments `blockedCells` when not
 *   self-closing, so its later close cannot end the canonical cell;
 * - other openings under a row-child/cell barrier return nothing and only track matching nesting;
 *   fake non-self-closing cells additionally update the blocked counter described above.
 */
export function pushRowOpen(
	state: WorksheetStructure,
	name: string,
	selfClosing: boolean,
): WorksheetOpen | undefined {
	const owner = state.rowChild;
	if (owner === undefined) {
		if (name === "c") {
			if (!selfClosing) state.rowChild = createRegionChild(name, "c");
			return "c";
		}
		if (!selfClosing) state.rowChild = createRegionChild(name, "other");
		return undefined;
	}
	if (owner.kind === "c" && name === "c") {
		if (state.cellBarrier !== undefined) {
			if (!selfClosing) state.blockedCells++;
			pushNestedOpen(state.cellBarrier, name, selfClosing);
			return undefined;
		}
		// A new cell recovers a missing prior </c>, matching the row assembler's existing contract.
		state.rowChild = selfClosing ? undefined : createRegionChild(name, "c");
		state.cellBarrier = undefined;
		state.cellChild = undefined;
		return "c";
	}
	if (owner.kind === "other") {
		if (name === "c" && !selfClosing) state.blockedCells++;
		pushNestedOpen(owner, name, selfClosing);
	} else if (state.cellBarrier !== undefined) {
		pushNestedOpen(state.cellBarrier, name, selfClosing);
		// Count the direct child's name inside the extension too, so that only the direct child's own
		// close, never a same-name element's from inside the extension, can end it there.
		if (state.cellChild !== undefined) pushNestedOpen(state.cellChild, name, selfClosing);
	} else if (!isCellContent(name)) {
		if (!selfClosing) state.cellBarrier = createRegionChild(name, "other");
	} else if (state.cellChild !== undefined) pushNestedOpen(state.cellChild, name, selfClosing);
	else return pushCellChildOpen(state, name, selfClosing);
	// Known elements beneath an active cell are its value/formula/rich-text content.
	return undefined;
}

/**
 * Process known cell content that opens directly beneath an admitted cell, outside any barrier.
 * Examples: `<v>` returns `"v"` and records the cell's direct child until `</v>`; `<f/>` (a shared
 * formula's dependent) returns `"f"` and records nothing; `<t>` (not a direct child the schema
 * allows) returns nothing and records an `"other"` direct child.
 */
export function pushCellChildOpen(
	state: WorksheetStructure,
	name: string,
	selfClosing: boolean,
): CellChild | undefined {
	const kind = name === "f" || name === "v" || name === "is" ? name : undefined;
	if (!selfClosing) state.cellChild = createRegionChild(name, kind ?? "other");
	return kind;
}

/**
 * Process one child opening inside canonical `<sheetData>`.
 * Examples:
 * - no owner + `<row/>` returns `"row"` with no retained owner;
 * - no owner + `<row>` returns `"row"` and records the active row;
 * - no owner + `<ext>` returns nothing and opens a barrier; `<ext/>` opens none;
 * - active row + `<c>` delegates to row processing and returns `"c"` when admitted;
 * - every other non-row opening under an active row also delegates to row processing, preserving
 *   its known-content, self-closing unknown, and extension-barrier outcomes;
 * - active row + another `<row>` returns `"row"`, resets cell state, and recovers a missing
 *   prior `</row>` when no barrier is active;
 * - a `<row>` inside a barrier returns nothing and increments `blockedRows` when not self-closing;
 * - other openings under a non-row owner return nothing and only track matching nesting; fake
 *   non-self-closing rows additionally update the blocked counter described above.
 */
export function pushSheetDataOpen(
	state: WorksheetStructure,
	name: string,
	selfClosing: boolean,
): WorksheetOpen | undefined {
	const owner = state.regionChild;
	if (owner === undefined) {
		if (name === "row") {
			if (!selfClosing) state.regionChild = createRegionChild(name, "row");
			return "row";
		}
		if (!selfClosing) state.regionChild = createRegionChild(name, "other");
		return undefined;
	}
	if (owner.kind === "row") {
		if (name === "row") {
			const barrier = state.rowChild?.kind === "other" ? state.rowChild : state.cellBarrier;
			if (barrier !== undefined) {
				if (!selfClosing) state.blockedRows++;
				pushNestedOpen(barrier, name, selfClosing);
				return undefined;
			}
			// A new row recovers a missing prior </row>; its cell scope starts clean.
			state.regionChild = selfClosing ? undefined : createRegionChild(name, "row");
			state.rowChild = undefined;
			state.cellBarrier = undefined;
			state.cellChild = undefined;
			state.blockedRows = 0;
			state.blockedCells = 0;
			return "row";
		}
		return pushRowOpen(state, name, selfClosing);
	}
	if (name === "row" && !selfClosing) state.blockedRows++;
	pushNestedOpen(owner, name, selfClosing);
	return undefined;
}

/**
 * Process one close beneath the active canonical region.
 * Examples:
 * - no region child or an unrelated close returns nothing and leaves owners unchanged;
 * - closing a cols child/barrier returns nothing; the matching outer close clears it, while a
 *   matching nested close only reduces its same-name depth;
 * - closing a blocked fake `row`/`c` returns nothing and decrements its fixed-size counter;
 * - closing the canonical row returns `"row"` and clears all row/cell scope, even across an
 *   unclosed lower extension barrier;
 * - closing the canonical cell returns `"c"` and clears cell scope, also across such a barrier;
 * - a close for a direct, row-child, or cell extension owner returns nothing; only its matching
 *   outer close clears that barrier, while inner same-name closes only reduce its depth;
 * - with no barrier open, closing the cell's direct child returns `"f"`, `"v"` or `"is"` for those
 *   (nothing for other content) and clears it; a nested same-name close only reduces its depth;
 * - inside a cell barrier, a close of the direct child's name ends the direct child, and the barrier
 *   with it, only when it is not a same-name element's from inside the barrier: an extension's
 *   `</x:v>` cannot close the cell's value, while the cell's own `</v>` recovers across an unclosed
 *   extension, as `</c>` and `</row>` do (F11.26).
 */
export function pushRegionChildClose(
	state: WorksheetStructure,
	name: string,
): WorksheetClose | undefined {
	const owner = state.regionChild;
	if (owner === undefined) return undefined;
	if (state.region === "cols") {
		if (closesOwner(owner, name)) {
			state.regionChild = undefined;
			state.rootSameNameDepth = 0;
		}
		return undefined;
	}
	if (name === "row" && state.blockedRows > 0) {
		state.blockedRows--;
		return undefined;
	}
	if (owner.kind === "row" && name === "row") {
		state.regionChild = undefined;
		state.rowChild = undefined;
		state.cellBarrier = undefined;
		state.cellChild = undefined;
		state.blockedRows = 0;
		state.blockedCells = 0;
		state.rootSameNameDepth = 0;
		return "row";
	}
	if (name === "c" && state.blockedCells > 0) {
		state.blockedCells--;
		return undefined;
	}
	if (owner.kind === "row" && state.rowChild?.kind === "c" && name === "c") {
		state.rowChild = undefined;
		state.cellBarrier = undefined;
		state.cellChild = undefined;
		state.blockedRows = 0;
		state.blockedCells = 0;
		state.rootSameNameDepth = 0;
		return "c";
	}
	if (owner.kind === "other") {
		if (closesOwner(owner, name)) {
			state.regionChild = undefined;
			state.blockedRows = 0;
			state.blockedCells = 0;
			state.rootSameNameDepth = 0;
		}
		return undefined;
	}
	if (state.rowChild?.kind === "other") {
		if (closesOwner(state.rowChild, name)) {
			state.rowChild = undefined;
			state.blockedRows = 0;
			state.blockedCells = 0;
			state.rootSameNameDepth = 0;
		}
		return undefined;
	}
	const child = state.cellChild;
	if (state.cellBarrier !== undefined && (child === undefined || name !== child.name)) {
		if (closesOwner(state.cellBarrier, name)) {
			state.cellBarrier = undefined;
			state.blockedRows = 0;
			state.blockedCells = 0;
			state.rootSameNameDepth = 0;
		}
		return undefined;
	}
	if (child !== undefined && closesOwner(child, name)) {
		state.cellChild = undefined;
		state.cellBarrier = undefined;
		state.blockedRows = 0;
		state.blockedCells = 0;
		return child.kind === "f" || child.kind === "v" || child.kind === "is"
			? child.kind
			: undefined;
	}
	return undefined;
}

/**
 * Release the current worksheet direct child and all nested structural state.
 * Example: an active `sheetData > row > c > ext` becomes no direct child, region, row, cell, cell
 * child, or barrier, with both blocked counters and both same-name depths reset to zero. The
 * worksheet phase itself is preserved. The function returns no value.
 */
export function clearDirectChild(state: WorksheetStructure): void {
	state.directChild = undefined;
	state.sameNameDepth = 0;
	state.region = undefined;
	state.regionChild = undefined;
	state.rowChild = undefined;
	state.cellBarrier = undefined;
	state.cellChild = undefined;
	state.blockedRows = 0;
	state.blockedCells = 0;
	state.rootSameNameDepth = 0;
}
