import { mimeForMediaPath, parseDrawing, parseRels, resolveTarget } from "../../ooxml";
import type { SheetImage } from "../../types";
import { writeXlsx } from "../../writer";
import { openZip } from "../../zip";
import { decodeText } from "../decode";
import { directoryOf, relsPathFor } from "../opc-path";

// F11.96 fixtures: workbooks with pictures, and the 1.0.0 per-sheet picture loader kept as an oracle.

/** The relationship type of an OPC drawing part. */
export const DRAWING_REL =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/**
 * A tiny PNG-signature payload, distinct per `n`, for picture bytes.
 *
 * @example pictureBytes(1) // Uint8Array [0x89, 0x50, 0x4e, 0x47, 1]
 */
export function pictureBytes(n: number): Uint8Array {
	return new Uint8Array([0x89, 0x50, 0x4e, 0x47, n]);
}

/**
 * A workbook with sheets S1, S2, … whose sheet `i` shows one picture with `pictures[i]` as bytes
 * (or none for `undefined`).
 *
 * @example await pictureWorkbook([pictureBytes(1), pictureBytes(1)]) // one media part used by two drawings
 */
export function pictureWorkbook(
	pictures: readonly (Uint8Array | undefined)[],
): Promise<Uint8Array> {
	return writeXlsx({
		sheets: pictures.map((bytes, i) => ({
			name: `S${i + 1}`,
			rows: [[i]],
			...(bytes === undefined
				? {}
				: {
						images: [
							{
								anchor: { from: { col: 1, row: 1 }, ext: { cx: 100, cy: 100 } },
								bytes,
								mime: "image/png",
							},
						],
					}),
		})),
	});
}

/**
 * The pictures the 1.0.0 loader returned for the sheet at `sheetPath`: every drawing relationship
 * read on its own, with a media cache per drawing. Kept only as an oracle for workbooks that name
 * each drawing and media part once.
 *
 * @example await legacySheetImages(xlsx, "xl/worksheets/sheet1.xml") // [{ anchor, bytes, mime }]
 * @example await legacySheetImages(xlsx, "xl/worksheets/no-such-sheet.xml") // []
 */
export async function legacySheetImages(
	bytes: Uint8Array,
	sheetPath: string,
): Promise<SheetImage[]> {
	const zip = openZip(bytes);
	const sheetRelsPath = relsPathFor(sheetPath);
	if (!zip.has(sheetRelsPath)) return [];
	const sheetRels = parseRels(decodeText(await zip.read(sheetRelsPath)));
	const images: SheetImage[] = [];
	for (const drawingRel of sheetRels.values()) {
		if (!drawingRel.type.endsWith("/drawing") || drawingRel.targetMode === "External") continue;
		const drawingPath = resolveTarget(directoryOf(sheetPath), drawingRel.target);
		if (!zip.has(drawingPath)) continue;
		const anchors = parseDrawing(decodeText(await zip.read(drawingPath)));
		const embedRelsPath = relsPathFor(drawingPath);
		if (anchors.length === 0 || !zip.has(embedRelsPath)) continue;
		const embedRels = parseRels(decodeText(await zip.read(embedRelsPath)));
		for (const anchor of anchors) {
			const rel = embedRels.get(anchor.embed);
			if (rel === undefined || rel.targetMode === "External") continue;
			const mediaPath = resolveTarget(directoryOf(drawingPath), rel.target);
			if (!zip.has(mediaPath)) continue;
			images.push({
				anchor: {
					from: anchor.from,
					...(anchor.to !== undefined ? { to: anchor.to } : {}),
					...(anchor.ext !== undefined ? { ext: anchor.ext } : {}),
					...(anchor.editAs !== undefined ? { editAs: anchor.editAs } : {}),
				},
				bytes: await zip.read(mediaPath),
				mime: mimeForMediaPath(mediaPath),
				...(anchor.name !== undefined ? { name: anchor.name } : {}),
			});
		}
	}
	return images;
}
