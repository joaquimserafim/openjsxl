import { XlsxError } from "../errors";
import { createMediaRegistry } from "./images";
import {
	contentTypesXml,
	encode,
	packageRelsXml,
	requireWorkbookObject,
	sheetSideParts,
	themeToEmit,
	validateDefinedNames,
	validateSheetMeta,
	workbookRelsXml,
	workbookXml,
} from "./parts";
import { createTableContext, streamWorksheet } from "./sheet";
import { type StreamPart, streamZip } from "./stream-zip";
import { createStyleRegistry } from "./styles";
import type { StreamSheetInput, StreamWorkbookInput, WriteOptions } from "./types";

// streamXlsx (F5.1) — the constant-memory sibling of writeXlsx. It emits the same OPC parts (built by
// the shared helpers in parts.ts), but each worksheet's rows flow through the streaming zip instead
// of buffering the sheet, so peak memory tracks one compression window rather than the whole file.
//
// Ordering is load-bearing: the worksheets stream FIRST, so styles intern as their rows flow; only
// once every sheet has drained are styles.xml / theme1.xml (which depend on what interned) known and
// emitted, followed by workbook.xml, the rels, and [Content_Types].xml. Zip part order is irrelevant
// to OPC — everything resolves through relationships — so a reader reassembles the workbook the same
// as from writeXlsx. Byte-identity with writeXlsx is NOT a goal (stream layout differs); equivalence
// is asserted through the reader.

// The streaming analogue of the buffered writer's Array.isArray(rows) check: rows must be iterable
// (a materialized array, or a sync/async iterable such as a DB cursor).
function isIterableRows(rows: unknown): boolean {
	if (rows === null || typeof rows !== "object") return false;
	const r = rows as { [Symbol.iterator]?: unknown; [Symbol.asyncIterator]?: unknown };
	return (
		typeof r[Symbol.iterator] === "function" || typeof r[Symbol.asyncIterator] === "function"
	);
}

async function* buildStreamParts(
	workbook: StreamWorkbookInput,
	options?: WriteOptions,
): AsyncGenerator<StreamPart> {
	// Reject a non-object workbook typed, before the property read that would otherwise raw-throw. Then
	// read the caller's `sheets` ONCE (single-read TOCTOU, matching writeXlsx), and validate and emit
	// from that single array and the resolved names/states.
	requireWorkbookObject(workbook);
	const sheets = workbook.sheets;
	const { states, names } = validateSheetMeta(sheets);
	for (let i = 0; i < sheets.length; i++) {
		if (!isIterableRows((sheets[i] as StreamSheetInput).rows)) {
			throw new XlsxError(
				"invalid-input",
				`sheet "${names[i]}": rows must be an iterable of row arrays`,
			);
		}
	}
	const date1904 = options?.date1904 === true;
	// Read the caller's optional carried theme ONCE (single-read TOCTOU, matching writeXlsx).
	const carriedTheme = workbook.themeXml;
	// Read + validate the optional defined names ONCE (F10.1), matching writeXlsx — absent → byte-identical.
	const definedNames = validateDefinedNames(workbook.definedNames, names.length);
	const styles = createStyleRegistry();
	const media = createMediaRegistry();
	const tableCtx = createTableContext();

	// Prepare every sheet upfront — this validates geometry/metadata and builds each header/footer +
	// rel/comment/drawing parts (interning any image bytes into `media`) — while the rows stay lazy
	// inside each `chunks` generator. So `media` is fully populated before any byte streams.
	const prepared = sheets.map((sheet, i) =>
		streamWorksheet(sheet, i, date1904, styles, media, tableCtx),
	);

	// Which sheets carry comments / drawings — derived upfront from the prepared results, so the
	// content-type map at the end knows without tracking it mid-loop.
	const commentSheets = prepared.flatMap((w, i) => (w.commentsXml !== undefined ? [i] : []));
	const drawingSheets = prepared.flatMap((w, i) => (w.drawingXml !== undefined ? [i] : []));
	const tablePartNumbers = prepared.flatMap((w) => (w.tables ?? []).map((t) => t.number));

	// Worksheets stream first, each followed by its own side parts (rels/comments/VML/drawing) — names
	// owned by sheetSideParts, shared with the buffered writer. streamZip consumes each part fully
	// before the next, so after this loop every sheet's rows have rendered and the registry is complete.
	for (let i = 0; i < prepared.length; i++) {
		const w = prepared[i] as (typeof prepared)[number];
		yield { name: `xl/worksheets/sheet${i + 1}.xml`, data: w.chunks };
		for (const part of sheetSideParts(i, w)) {
			yield { name: part.name, data: encode(part.xml) };
		}
	}
	// Workbook-level media parts (binary, already deduped during preparation).
	for (const part of media.parts()) yield { name: part.name, data: part.data };

	// Now the registry is final, so the style/theme parts and the content-type map are known.
	const needStyles = styles.needed();
	const needTheme = styles.usesTheme();
	if (needStyles) yield { name: "xl/styles.xml", data: encode(styles.stylesXml()) };
	if (needTheme) yield { name: "xl/theme/theme1.xml", data: encode(themeToEmit(carriedTheme)) };
	yield {
		name: "xl/workbook.xml",
		data: encode(workbookXml(names, states, date1904, definedNames)),
	};
	yield {
		name: "xl/_rels/workbook.xml.rels",
		data: encode(workbookRelsXml(sheets.length, needStyles, needTheme)),
	};
	yield {
		name: "[Content_Types].xml",
		data: encode(
			contentTypesXml(
				sheets.length,
				needStyles,
				needTheme,
				commentSheets,
				drawingSheets,
				media.extensions(),
				tablePartNumbers,
			),
		),
	};
	yield { name: "_rels/.rels", data: encode(packageRelsXml()) };
}

/**
 * Serialize a workbook to `.xlsx` as a {@link ReadableStream} of bytes with roughly constant memory —
 * the writer mirror of the reader's `streamSheetRows` (F5.1). Each sheet's `rows` may be a sync or
 * async iterable (an array, a generator, a DB cursor), pulled only as the consumer reads the output,
 * so a slow source is never outpaced. The bytes read back through {@link openXlsx} with the same
 * values, types, styles, geometry, and metadata as {@link writeXlsx} — a streamed sheet just omits
 * the optional `<dimension>` (its bounds aren't known upfront). Invalid input surfaces as an
 * {@link XlsxError} on the stream (the read rejects), not synchronously.
 */
export function streamXlsx(
	workbook: StreamWorkbookInput,
	options?: WriteOptions,
): ReadableStream<Uint8Array> {
	const gen = streamZip(buildStreamParts(workbook, options));
	return new ReadableStream<Uint8Array>({
		async pull(controller) {
			try {
				const { done, value } = await gen.next();
				if (done) controller.close();
				else controller.enqueue(value);
			} catch (err) {
				controller.error(err);
			}
		},
		async cancel() {
			// The consumer went away — let the row sources release (close DB cursors, etc.).
			await gen.return?.(undefined);
		},
	});
}
