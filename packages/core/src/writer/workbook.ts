import { XlsxError } from "../errors";
import type { SheetState } from "../types";
import { createMediaRegistry } from "./images";
import {
	contentTypesXml,
	encode,
	packageRelsXml,
	sheetSideParts,
	themeToEmit,
	validateSheetMeta,
	workbookRelsXml,
	workbookXml,
} from "./parts";
import { createTableContext, worksheetXml } from "./sheet";
import { createStyleRegistry } from "./styles";
import type { SheetInput, WorkbookInput, WriteOptions } from "./types";
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

// Validate the workbook's sheets (already read once by the caller): shared name/state checks
// (parts.ts) plus the buffered writer's own requirement that each sheet's `rows` is a materialized
// array (the streaming writer accepts an iterable instead). Returns the resolved names + visibility
// states for single-read emission.
function validate(sheets: readonly SheetInput[]): { states: SheetState[]; names: string[] } {
	const { states, names } = validateSheetMeta(sheets);
	for (let i = 0; i < sheets.length; i++) {
		if (!Array.isArray((sheets[i] as SheetInput).rows)) {
			throw new XlsxError("invalid-input", `sheet "${names[i]}": rows must be an array`);
		}
	}
	return { states, names };
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
	// Read the caller's `sheets` ONCE (a getter/Proxy must not vary it between validation and
	// emission), then validate that single array and emit from it and the resolved names/states.
	const sheets = workbook.sheets;
	const { states, names } = validate(sheets);
	const date1904 = options?.date1904 === true;
	// Read the caller's optional carried theme ONCE too. It only matters if a style needs a theme part.
	const carriedTheme = workbook.themeXml;

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
		worksheetXml(sheet, i, date1904, styles, media, tableCtx),
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
	add("xl/workbook.xml", workbookXml(names, states, date1904));
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
