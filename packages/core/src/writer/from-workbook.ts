import { XlsxError } from "../errors";
import { type CellRef, formatRef, MAX_COL, MAX_ROW, parseRef } from "../ooxml/a1";
import { MAX_TABLE_NAME_LEN } from "../ooxml/table";
import type { DefinedName } from "../ooxml/workbook";
import type { Workbook } from "../reader/workbook";
import type {
	Cell,
	ColumnProps,
	Comment,
	ConditionalFormatting,
	DataValidation,
	FreezePane,
	Hyperlink,
	RowProps,
	SheetAutoFilter,
	SheetImage,
	SheetProtection,
	SheetState,
	TableInfo,
	Worksheet,
} from "../types";
import type { CellInput, SheetInput, WorkbookInput } from "./types";

/**
 * Return `name` if unseen (case-insensitively), else a suffixed variant (`name_2`, `name_3`, …) that
 * stays within {@link MAX_TABLE_NAME_LEN}; records whatever it returns in `seen`. A suffixed legal name
 * stays legal (appends `_<digits>` — no whitespace, no cell-ref shape, letter-first preserved).
 * Exported for unit testing; the bridge uses it so the reader's F9.5 name normalization (which can map
 * two distinct illegal names to one legal string) can't produce a duplicate the writer rejects.
 */
export function uniquifyTableName(name: string, seen: Set<string>): string {
	if (!seen.has(name.toLowerCase())) {
		seen.add(name.toLowerCase());
		return name;
	}
	for (let n = 2; ; n++) {
		const suffix = `_${n}`;
		let head = name.slice(0, MAX_TABLE_NAME_LEN - suffix.length);
		// `slice` cuts at UTF-16 code units, so a truncation landing inside an astral character
		// leaves a lone high surrogate — an ILLEGAL name that would abort the very write this dedup
		// exists to protect (F9.6). Drop the half pair so the candidate is always legal.
		const last = head.charCodeAt(head.length - 1);
		if (last >= 0xd800 && last <= 0xdbff) head = head.slice(0, -1);
		const candidate = `${head}${suffix}`;
		if (!seen.has(candidate.toLowerCase())) {
			seen.add(candidate.toLowerCase());
			return candidate;
		}
	}
}

/**
 * Drop later per-scope case-insensitive DUPLICATE defined names (F10.1). The tolerant reader returns
 * each name individually writer-legal but does NOT dedupe across the workbook, so a foreign file with
 * two same-scope names that collide case-insensitively (which Excel itself forbids) would otherwise
 * make the subsequent writeXlsx reject the pair and abort a legitimate save. Keep the FIRST occurrence
 * in document order — a defined name can't be suffix-renamed like a table (formulas reference it by
 * name), so a genuine collision is resolved by dropping, exactly as Excel repairs such a file.
 * Exported for unit testing. Returns the input array UNCHANGED (same reference) when there are no
 * duplicates, so the common clean-file path adds no allocation.
 */
export function dedupeDefinedNames(names: readonly DefinedName[]): readonly DefinedName[] {
	const seen = new Set<string>();
	let duplicate = false;
	for (const dn of names) {
		const scopeKey = dn.localSheetId !== undefined ? String(dn.localSheetId) : "*";
		const key = `${scopeKey} ${dn.name.toUpperCase()}`;
		if (seen.has(key)) {
			duplicate = true;
			break;
		}
		seen.add(key);
	}
	if (!duplicate) return names;
	seen.clear();
	const out: DefinedName[] = [];
	for (const dn of names) {
		const scopeKey = dn.localSheetId !== undefined ? String(dn.localSheetId) : "*";
		const key = `${scopeKey} ${dn.name.toUpperCase()}`;
		if (seen.has(key)) continue;
		seen.add(key);
		out.push(dn);
	}
	return out;
}

// Bridge the reader to the writer: turn an open Workbook into the plain-data input writeXlsx wants,
// so a file can be read, optionally tweaked, and written back. This closes the round trip (F3.3);
// since F4.4 it carries STYLES too — the CellStyle that sheet.style(ref) resolves is exactly the
// shape writeXlsx accepts, so styles cross the bridge as a structural pass-through.
//
// Fidelity is scoped to values, types, sheet names/order, and cell styles (number format, font,
// fill, border, alignment). What the writer can't represent is NOT carried across and is
// documented, not silently mangled:
//   - error cells (no formula) → written as their literal text (e.g. "#DIV/0!"), so they become strings
//   (column widths, row heights, hidden flags, frozen panes — F4.5 — merges, hyperlinks, sheet
//   visibility — F4.6 — cell comments — F5.2 — formula TEXT — F5.4 — and anchored PICTURES — F6.4 —
//   DO carry across)
//   Picture degradations (documented; the reader already skips these on read, so they never reach
//   the bridge): absolute-anchored pictures and non-picture drawing objects (shapes, charts) are
//   dropped at read. A picture whose media type is outside the writer's allowlist (anything but
//   the MEDIA_MIME_TO_EXT set — png, jpeg, gif, bmp, tiff, webp, x-emf, x-wmf) makes the
//   subsequent writeXlsx throw a TYPED invalid-input — the whole rewrite refuses rather than
//   silently dropping the picture.
//   NOW CARRIED (F10.1): DEFINED NAMES — Workbook.definedNames crosses the bridge into
//   WorkbookInput.definedNames, so a named-range workbook keeps its <definedNames> and its formulas
//   resolve on open instead of recalculating to #NAME?. The reader already dropped any name it could
//   not re-emit (an illegal identifier, an empty/oversized refersTo, a sheet-scope past the sheet
//   list), and the bridge drops a later per-scope case-insensitive DUPLICATE (a name can't be
//   suffix-renamed like a table — formulas reference it), so what crosses is always writer-legal.
//   NOT carried (documented drops): IN-CELL RICH TEXT — per-run formatting flattens to plain text at
//   read (the concatenated value is what crosses the bridge).
//   Formula degradations (documented, values stay exact): a SHARED-formula dependent carries its
//   TRANSLATED text (not the shared grouping); an ARRAY formula carries the master's text as a plain
//   formula (the spilled cells keep their cached values); a `dataTable` formula carries no text (its
//   cached value survives); an error-typed cached value on a formula cell writes as its string text.
//   - locale-only number formats (ids 23–36, 50–58) → no portable code exists; the format flattens
//     (a date keeps its value and date-ness via the implicit format)
// Three documented FLATTENINGS (values stay exact; the file's internal spelling normalizes):
//   - a tolerated NON-CANONICAL cell ref spelling (e.g. lowercase "a1") re-emits canonically
//     ("A1") — same grid slot, same value/type/style; A1 notation is case-insensitive and the
//     writer's input is positional, so canonical is the only possible emission.
//   - row/column DEFAULT styles resolve into per-cell styles (the effective style each cell already
//     shows through style(ref)) — the rewritten file styles cells directly instead of via defaults.
// (Custom themes NO LONGER flatten — since F5.3 the source theme1.xml is carried verbatim, so
//  {theme, tint} indexes re-render against the SAME colors on rewrite.)
// A source sheet name the writer rejects (>31 chars, forbidden characters, duplicate) will make the
// subsequent writeXlsx throw invalid-input — such a workbook isn't Excel-valid to re-emit as-is.
// Likewise a cell whose `r` attribute is unaddressable garbage (the tolerant reader keeps it,
// faithful to the document, e.g. a 300-letter column ref): it has no grid position to re-write to,
// so the bridge throws a TYPED invalid-input instead of silently dropping it or crashing bare.

// The reader already exposes each cell as its JS value: null for empty, and string / number /
// boolean / Date for the typed kinds. The lone nuance is 'error', whose value is the error text —
// the writer has no error-cell form, so it round-trips as that string. A cell with a resolved
// style wraps into the writer's { value, style } form — including a styled EMPTY cell (<c s/>),
// which is how a border or fill on a blank cell survives the trip. style(ref) returns one cached
// object per distinct format, so the wrapping adds O(distinct formats) objects, not O(cells).
function cellToInput(worksheet: Worksheet, cell: Cell): CellInput {
	const style = worksheet.style(cell.ref);
	const formula = worksheet.formula(cell.ref);
	if (formula !== undefined) {
		// A formula cell (F5.4) carries the formula text plus its cached value (and any style), so a
		// non-recalculating consumer still sees the last computed result. An error-typed cached value
		// writes as its string text — the formula is what matters; Excel recomputes the real error.
		return style === undefined
			? { formula, value: cell.value }
			: { formula, value: cell.value, style };
	}
	return style === undefined ? cell.value : { value: cell.value, style };
}

/**
 * Convert an open {@link Workbook} into {@link WorkbookInput} for {@link writeXlsx}. Each populated
 * cell is placed at its own A1 reference with its resolved style (if any), preserving sheet names
 * and tab order. Rows/columns are left sparse (array holes) — the writer treats a hole as an empty
 * cell — so a workbook with a few far-apart cells does not materialize a dense grid. An unstyled
 * workbook yields bare values only: the exact pre-styles WorkbookInput, and the exact same bytes
 * on rewrite.
 */
export async function workbookToInput(workbook: Workbook): Promise<WorkbookInput> {
	const sheets: SheetInput[] = [];
	// Table names must be workbook-unique (case-insensitively) or the writer rejects them. The tolerant
	// reader's name normalization (F9.5) can map two DISTINCT illegal names to the SAME legal string, so
	// dedupe here — suffixing a collision — instead of aborting the whole save on the duplicate.
	const seenTableNames = new Set<string>();
	for (const info of workbook.sheets) {
		const worksheet = workbook.sheet(info.name);
		const rows: CellInput[][] = [];
		// Verbatim source ref per occupied grid slot (canonical ref → the ref string that filled
		// it). The reader's cell identity is the VERBATIM ref: "A1" and "a1" are two different
		// cells to cell(), yet parse to one grid slot — silently letting one overwrite the other
		// would vanish a value with no error. Same-spelling duplicates stay last-wins, which is
		// exactly how the reader's own cell() map resolves them, so the two sides agree.
		const occupied = new Map<string, string>();
		for await (const row of worksheet.rows()) {
			for (const cell of row.cells) {
				// Place by the cell's own ref, not the row index — the two agree for well-formed
				// files, and the ref is authoritative for the column in either case.
				let placed: CellRef | undefined;
				try {
					placed = parseRef(cell.ref);
				} catch {
					placed = undefined;
				}
				// A ref that doesn't parse OR lies outside Excel's grid has no writable position.
				// The grid cap also matters mechanically: `rows` is indexed by row number, so a
				// hostile row like 1e14 would otherwise become an array LENGTH the writer iterates.
				if (placed === undefined || placed.row > MAX_ROW || placed.col > MAX_COL) {
					const shown = cell.ref.length > 24 ? `${cell.ref.slice(0, 24)}…` : cell.ref;
					throw new XlsxError(
						"invalid-input",
						`sheet "${info.name}": cell reference "${shown}" has no writable grid position`,
					);
				}
				const canonical = formatRef(placed);
				const prior = occupied.get(canonical);
				if (prior !== undefined && prior !== cell.ref) {
					throw new XlsxError(
						"invalid-input",
						`sheet "${info.name}": cells "${prior}" and "${cell.ref}" are distinct to the reader but occupy one grid position (${canonical})`,
					);
				}
				occupied.set(canonical, cell.ref);
				const { col, row: rowNum } = placed;
				let rowArr = rows[rowNum - 1];
				if (rowArr === undefined) {
					rowArr = [];
					rows[rowNum - 1] = rowArr;
				}
				rowArr[col - 1] = cellToInput(worksheet, cell);
			}
		}
		// Geometry (F4.5) and structural metadata (F4.6): carried only when present, so a workbook
		// using neither produces the exact same WorkbookInput — and the exact same bytes — as before.
		const sheet: {
			name: string;
			rows: CellInput[][];
			columns?: readonly ColumnProps[];
			rowProperties?: Readonly<Record<number, RowProps>>;
			freeze?: FreezePane;
			merges?: readonly string[];
			hyperlinks?: readonly Hyperlink[];
			state?: SheetState;
			comments?: readonly Comment[];
			images?: readonly SheetImage[];
			tables?: readonly TableInfo[];
			dataValidations?: readonly DataValidation[];
			conditionalFormatting?: readonly ConditionalFormatting[];
			autoFilter?: SheetAutoFilter;
			protection?: SheetProtection;
		} = { name: info.name, rows };
		const columns = worksheet.columns;
		if (columns.length > 0) sheet.columns = columns;
		const rowProperties = worksheet.rowProperties;
		if (rowProperties.size > 0) sheet.rowProperties = Object.fromEntries(rowProperties);
		const freeze = worksheet.freeze;
		if (freeze !== undefined) sheet.freeze = freeze;
		const merges = worksheet.mergedCells;
		if (merges.length > 0) sheet.merges = merges;
		const hyperlinks = worksheet.hyperlinks;
		if (hyperlinks.length > 0) sheet.hyperlinks = hyperlinks;
		if (info.state !== "visible") sheet.state = info.state;
		const comments = worksheet.comments;
		if (comments.length > 0) sheet.comments = comments;
		// Pictures (F6.4). `images()` is async (media needs decompression) and degrading — it returns
		// only cell-anchored pictures with resolvable bytes, so what it hands back is exactly a valid
		// writer input. Attach only when non-empty: an imageless workbook keeps the byte-identity path.
		const images = await worksheet.images();
		if (images.length > 0) sheet.images = images;
		// Tables (F9.1) — structural pass-through; the writer re-derives column names from the header
		// row (which the bridge also carries), so a read table rewrites cleanly. Names are deduped
		// workbook-wide (F9.5) — a unique name is passed through UNCHANGED (byte-identity), a collision
		// gets a suffix.
		const tables = worksheet.tables;
		if (tables.length > 0) {
			sheet.tables = tables.map((t) => {
				const unique = uniquifyTableName(t.name, seenTableNames);
				return unique === t.name ? t : { ...t, name: unique };
			});
		}
		// Data validations (F9.2) — structural pass-through; the reader's rules ARE writer input.
		const dataValidations = worksheet.dataValidations;
		if (dataValidations.length > 0) sheet.dataValidations = dataValidations;
		// Conditional formatting (F9.3) — structural pass-through; the inline dxf re-interns on write.
		const conditionalFormatting = worksheet.conditionalFormatting;
		if (conditionalFormatting.length > 0) sheet.conditionalFormatting = conditionalFormatting;
		// autoFilter (F10.2) — structural pass-through. The writer re-emits <autoFilter> and re-synthesizes
		// the paired _xlnm._FilterDatabase name (which the reader stripped), so the filter round-trips once.
		const autoFilter = worksheet.autoFilter;
		if (autoFilter !== undefined) sheet.autoFilter = autoFilter;
		// Sheet protection (F10.3) — structural pass-through; per-cell locked/hidden rides along on each
		// cell's carried style. Password material is carried verbatim (never recomputed).
		const protection = worksheet.protection;
		if (protection !== undefined) sheet.protection = protection;
		sheets.push(sheet);
	}
	// Carry the source theme verbatim (F5.3) so custom theme colors survive the rewrite. Absent when
	// the workbook has no theme part — then the writer falls back to the built-in Office theme.
	const themeXml = workbook.themeXml;
	// Carry defined names (F10.1) — each already writer-legal (the reader dropped what it couldn't
	// re-emit); dedupe a foreign file's per-scope duplicate so the save doesn't abort. Attach only when
	// non-empty, so a names-free workbook yields the exact same WorkbookInput — and bytes — as before.
	const definedNames = dedupeDefinedNames(workbook.definedNames);
	// Workbook-level protection (F10.3) — structural pass-through, carried verbatim.
	const protection = workbook.protection;
	const out: WorkbookInput = {
		sheets,
		...(themeXml !== undefined ? { themeXml } : {}),
		...(definedNames.length > 0 ? { definedNames } : {}),
		...(protection !== undefined ? { protection } : {}),
	};
	return out;
}
