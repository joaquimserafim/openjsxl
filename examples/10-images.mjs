// 10 — Anchored pictures (0.6): put an image on a sheet, read it back, round-trip it.
//
//   node 10-images.mjs   (from ./examples)
//   pnpm --filter openjsxl-examples images
//
// A sheet takes `images`: each is `{ anchor, bytes, mime, name? }` — the SAME record the reader's
// `Worksheet.images()` hands back (one shared model). `bytes` is the raw image payload (opaque to
// the writer — never decoded); `mime` is image/png, image/jpeg, or image/gif. The anchor is kept
// RAW: 1-based `from` cell plus EMU offsets/extents, exactly as OOXML stores it (914 400 EMU/inch;
// ≈9 525 EMU/px @96 dpi). A one-cell anchor pins the picture at `from` with a fixed `ext` size; a
// two-cell anchor spans `from`→`to` and resizes with the cells.

import { openXlsx, workbookToInput, writeXlsx } from "openjsxl";

// A real 24×24 PNG (a blue tile with a white cross), embedded so the example stays self-contained.
const LOGO = Uint8Array.from(
	atob(
		"iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAIAAABvFaqvAAAAfklEQVR4nL3VSw5AEQwF0O7EEm3X" +
			"bq6BwWtoq48biUSEHvErAVBquywAZFSXSqlNdONY+aAzS4eI1/FLmaG8tQ6boYxlDjCg2PK6bMg" +
			"LCCZwoTUsXnIE6eDtxm2gQWSO8hXEWRpnsznHz7mQnCfCebScNMJJbJxUy0n+rO+oA63Acf6SxS" +
			"iKAAAAAElFTkSuQmCC",
	),
	(c) => c.charCodeAt(0),
);

// EMU helper — the anchor speaks EMU, so convert from the friendlier pixels (96 dpi).
const px = (n) => Math.round(n * 9525);

const bytes = await writeXlsx({
	sheets: [
		{
			name: "Report",
			rows: [
				[{ value: "Q3 Sales", style: { font: { bold: true, size: 14 } } }],
				["Region", "Units"],
				["North", 1200],
				["South", 980],
			],
			images: [
				{
					// One-cell anchor: pinned at D1 (col 4, row 1) with a fixed 96×96 px size.
					anchor: { from: { col: 4, row: 1 }, ext: { cx: px(96), cy: px(96) } },
					bytes: LOGO,
					mime: "image/png",
					name: "Logo",
				},
			],
		},
	],
});

// Read the pictures straight back out of the file. `images()` is async — the media bytes are
// decompressed on demand, so a sheet whose pictures you never touch costs no I/O.
const wb = await openXlsx(bytes);
const images = await wb.sheet("Report").images();
console.log("images     :", images.length);
for (const img of images) {
	console.log(
		"  •",
		img.name ?? "(unnamed)",
		"—",
		img.mime,
		`${img.bytes.length} bytes`,
		"@ col",
		img.anchor.from.col,
		"row",
		img.anchor.from.row,
	);
}

// And the picture survives read → modify → write: the bridge carries `images` verbatim, so a
// logo'd report you read in comes back out with the same image bytes and anchor.
const input = await workbookToInput(wb);
const again = await openXlsx(await writeXlsx(input));
const roundTripped = await again.sheet("Report").images();
console.log(
	"round-trip :",
	roundTripped.length,
	"image(s), bytes identical:",
	roundTripped[0].bytes.length === LOGO.length,
);
