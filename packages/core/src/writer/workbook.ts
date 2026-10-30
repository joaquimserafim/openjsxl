import { XlsxError } from "../errors";
import type { SheetState } from "../types";
import { filterDatabaseName, validateDefinedNames } from "./defined-names";
import { createMediaRegistry } from "./images";
import {
	contentTypesXml,
	encode,
	optionsDate1904,
	packageRelsXml,
	requireWorkbookObject,
	sheetSideParts,
	themeToEmit,
	validateSheetMeta,
	workbookRelsXml,
	workbookXml,
} from "./parts";
import { createTableContext, worksheetXml } from "./sheet";
import { createStyleRegistry } from "./styles";
import type { SheetInput, WorkbookInput, WriteOptions } from "./types";
import { ownValue } from "./xml";
import { writeZip, type ZipInput } from "./zip";

// writeXlsx — the public workbook writer (F3.2, styled cells F4.2). Given a workbook described as
// plain data, emit the minimal valid OPC part set and pack it with writeZip. The output re-reads
// through openXlsx with the same values, types, sheet order, and styles, and opens in
// Excel/LibreOffice without a "repair" prompt.
//
// Minimal part set (only what a valid spreadsheet requires):
//   [Content_Types].xml            — content-type map for every part
//   _rels/.rels                    — package → xl/workbook.xml
//   xl/workbook.xml                — sheet list (names, ids, r:id links)
//   xl/_rels/workbook.xml.rels     — workbook → each worksheet (+ styles/theme when present)
//   xl/worksheets/sheetN.xml       — one per sheet
//   xl/styles.xml                  — ONLY when some cell needs a non-default format (a date, or
//                                    any F4.2 style); built by the shared StyleRegistry
//   xl/theme/theme1.xml            — ONLY when a written style uses a theme color; Excel resolves
//                                    {theme, tint} indexes against this part
//
// Strings are written inline (t="inlineStr"), so there is no sharedStrings.xml to buffer — the
// "value extractor" trade-off: a touch more bytes on disk for a single streaming pass and no table.
// The per-part XML builders live in parts.ts, shared verbatim with the streaming writer (F5.1).

// Cache each own sheet-array item on first access. Validation may stop at an earlier bad sheet;
// do not invoke a later caller getter until the validator actually reaches that item.
export function snapshotSheets<T>(input: readonly T[] | undefined): T[] {
	if (!Array.isArray(input)) {
		throw new XlsxError("invalid-input", "a workbook needs at least one sheet");
	}
	const source = input;
	const length = source.length;
	if (!Number.isInteger(length) || length <= 0 || length > 0xffffffff) {
		throw new XlsxError("invalid-input", "a workbook needs at least one sheet");
	}
	const sheets: T[] = [];
	function prepare(i: number): void {
		Object.defineProperty(sheets, i, {
			get() {
				const value = ownValue(source, i);
				// An absent index stays undefined for validation. Defining the cached value directly
				// also avoids inherited array setters and prevents later reads of the caller getter.
				Object.defineProperty(sheets, i, {
					value,
					writable: true,
					enumerable: true,
					configurable: true,
				});
				// Prepare only the next item. A huge sparse array must fail at its first absent
				// sheet without allocating one getter for every trailing hole.
				if (i + 1 < length) prepare(i + 1);
				return value;
			},
			enumerable: true,
			configurable: true,
		});
	}
	prepare(0);
	// The accessor puts this array in sparse storage before extending its length. new Array(length)
	// can eagerly allocate millions of unused slots before the first sheet is even validated.
	sheets.length = length;
	return sheets;
}

// Validate names/states and the buffered writer's materialized rows, retaining each row array.
function validate(sheets: readonly SheetInput[]): {
	states: SheetState[];
	names: string[];
	rows: SheetInput["rows"][];
} {
	const { states, names } = validateSheetMeta(sheets);
	const rows: SheetInput["rows"][] = [];
	for (let i = 0; i < sheets.length; i++) {
		const sheetRows = ownValue(sheets[i] as SheetInput, "rows");
		if (!Array.isArray(sheetRows)) {
			throw new XlsxError("invalid-input", `sheet "${names[i]}": rows must be an array`);
		}
		rows.push(sheetRows);
	}
	return { states, names, rows };
}

/**
 * Serialize a workbook to `.xlsx` bytes. Async because compression runs on the platform's
 * CompressionStream. Throws {@link XlsxError} with code `invalid-input` for anything that can't be
 * represented — no sheets, a bad sheet name, or a cell value that isn't a string, finite number,
 * boolean, Date, or null.
 */
export async function writeXlsx(
	workbook: WorkbookInput,
	options?: WriteOptions,
): Promise<Uint8Array> {
	// Reject a non-object workbook typed, before the property read that would otherwise raw-throw. Then
	// read the caller's `sheets` ONCE (a getter/Proxy must not vary it between validation and emission),
	// validate that single array, and emit from it and the resolved names/states.
	requireWorkbookObject(workbook);
	const sheets = snapshotSheets(ownValue(workbook, "sheets"));
	const { states, names, rows } = validate(sheets);
	// The date system: an explicit option wins, else the input's own flag (what the bridge carries
	// from Workbook.date1904 — F11.10), else the 1900 default. Read the input flag ONCE (TOCTOU).
	const inputDate1904 = ownValue(workbook, "date1904");
	if (inputDate1904 !== undefined && typeof inputDate1904 !== "boolean") {
		throw new XlsxError("invalid-input", "workbook.date1904 must be a boolean when present");
	}
	const optionDate1904 = optionsDate1904(options);
	const date1904 = optionDate1904 ?? inputDate1904 ?? false;
	// Read the caller's optional carried theme ONCE too. It only matters if a style needs a theme part.
	const carriedTheme = ownValue(workbook, "themeXml");
	// Read the caller's optional defined names ONCE (single-read TOCTOU) and validate into the trusted
	// array workbook.xml emits (F10.1). Absent → [] → no <definedNames> → byte-identical output.
	const definedNames = validateDefinedNames(ownValue(workbook, "definedNames"), sheets.length);
	// Read the caller's optional workbook protection ONCE (single-read TOCTOU); validated + emitted by
	// workbookXml. Absent → no <workbookProtection> → byte-identical output.
	const protection = ownValue(workbook, "protection");

	// Render every worksheet up front — this is also where invalid cell values and styles surface —
	// interning every style into one shared registry. What the registry saw decides whether
	// styles.xml (any non-default format) and theme1.xml (any theme color) are emitted at all.
	const styles = createStyleRegistry();
	// Pictures (F6.3) intern their bytes into one workbook-level media registry as sheets render, so
	// identical images are written as a single shared media part.
	const media = createMediaRegistry();
	// Tables (F9.1) get workbook-global ids/part numbers and workbook-wide unique names, so one shared
	// context threads through every sheet.
	const tableCtx = createTableContext();
	const worksheets = sheets.map((sheet, i) =>
		// One captured row array per validated sheet, in the same order.
		worksheetXml(
			sheet,
			rows[i] as SheetInput["rows"],
			i,
			date1904,
			styles,
			media,
			tableCtx,
			names[i] as string,
		),
	);
	const needStyles = styles.needed();
	const needTheme = styles.usesTheme();

	// Comments (F5.2) ride on a legacy VML drawing part; images (F6.3) on a drawingML part. The `vml`
	// Default / image Defaults + drawing Overrides are emitted only when those parts exist — all
	// derived inside contentTypesXml, so an imageless/commentless workbook stays byte-identical.
	const commentSheets = worksheets.flatMap((w, i) => (w.commentsXml !== undefined ? [i] : []));
	const drawingSheets = worksheets.flatMap((w, i) => (w.drawingXml !== undefined ? [i] : []));
	// Every table part number across the workbook, for the content-type Overrides.
	const tablePartNumbers = worksheets.flatMap((w) => (w.tables ?? []).map((t) => t.number));

	// Append the hidden _xlnm._FilterDatabase name for every sheet that declared an autoFilter (F10.2).
	// The sheet writer validated each range once (reused here) and rejects a caller-supplied _FilterDatabase,
	// so these can neither duplicate nor collide with the caller's names.
	const allDefinedNames = [
		...definedNames,
		...worksheets.flatMap((w, i) => {
			const name = names[i];
			return w.autoFilterRef !== undefined && name !== undefined
				? [filterDatabaseName(name, i, w.autoFilterRef)]
				: [];
		}),
	];

	const parts: ZipInput[] = [];
	const add = (name: string, xml: string): void => {
		parts.push({ name, data: encode(xml) });
	};

	add(
		"[Content_Types].xml",
		contentTypesXml(
			sheets.length,
			needStyles,
			needTheme,
			commentSheets,
			drawingSheets,
			media.extensions(),
			tablePartNumbers,
		),
	);
	add("_rels/.rels", packageRelsXml());
	add("xl/workbook.xml", workbookXml(names, states, date1904, allDefinedNames, protection));
	add("xl/_rels/workbook.xml.rels", workbookRelsXml(sheets.length, needStyles, needTheme));
	if (needStyles) add("xl/styles.xml", styles.stylesXml());
	if (needTheme) add("xl/theme/theme1.xml", themeToEmit(carriedTheme));

	// The worksheet body, then its side parts (rels/comments/VML/drawing) — names owned by sheetSideParts.
	worksheets.forEach((w, i) => {
		add(`xl/worksheets/sheet${i + 1}.xml`, w.xml);
		for (const part of sheetSideParts(i, w)) add(part.name, part.xml);
	});
	// Workbook-level media parts (binary, already deduped) — appended after the sheet parts.
	for (const part of media.parts()) parts.push({ name: part.name, data: part.data });

	return writeZip(parts);
}
