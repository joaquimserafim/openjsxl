import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import type { SheetImage } from "../../types";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F6.3 — picture write. Bytes are opaque to the writer (never decoded), so these use short arbitrary
// buffers; the mime → media extension → re-read mime chain round-trips regardless of content.

const PNG = (n: number): Uint8Array => new Uint8Array([0x89, 0x50, 0x4e, 0x47, n]);
const oneCell = (
	col: number,
	row: number,
	bytes: Uint8Array,
	extra?: Partial<SheetImage>,
): SheetImage => ({
	anchor: { from: { col, row }, ext: { cx: 100, cy: 100 } },
	bytes,
	mime: "image/png",
	...extra,
});

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		total += value.length;
	}
	const out = new Uint8Array(total);
	let o = 0;
	for (const c of chunks) {
		out.set(c, o);
		o += c.length;
	}
	return out;
}

const entryNames = (bytes: Uint8Array): string[] => [...openZip(bytes).entries.keys()].sort();

describe("writeXlsx — picture write (F6.3)", () => {
	it("round-trips one-cell and two-cell anchors through the reader", async () => {
		const png = PNG(1);
		const jpg = new Uint8Array([0xff, 0xd8, 0xff, 2]);
		const input: WorkbookInput = {
			sheets: [
				{
					name: "Pics",
					rows: [["hi"]],
					images: [
						{
							anchor: {
								from: { col: 2, row: 3, colOff: 9525, rowOff: 19050 },
								ext: { cx: 762000, cy: 571500 },
							},
							bytes: png,
							mime: "image/png",
							name: "Logo",
						},
						{
							anchor: {
								from: { col: 4, row: 4 },
								to: { col: 6, row: 8 },
								editAs: "oneCell",
							},
							bytes: jpg,
							mime: "image/jpeg",
						},
					],
				},
			],
		};
		const wb = await openXlsx(await writeXlsx(input));
		const images = await wb.sheet("Pics").images();
		expect(images).toEqual([
			{
				anchor: {
					from: { col: 2, row: 3, colOff: 9525, rowOff: 19050 },
					ext: { cx: 762000, cy: 571500 },
				},
				bytes: png,
				mime: "image/png",
				name: "Logo",
			},
			{
				anchor: {
					from: { col: 4, row: 4, colOff: 0, rowOff: 0 },
					to: { col: 6, row: 8, colOff: 0, rowOff: 0 },
					editAs: "oneCell",
				},
				bytes: jpg,
				mime: "image/jpeg",
				name: "Image 2", // defaulted deterministically
			},
		]);
	});

	it("deduplicates identical bytes into one media part (near-misses stay separate)", async () => {
		const a = PNG(1);
		const aAgain = PNG(1); // equal bytes, different buffer
		const b = PNG(2); // same length, one byte different
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[1]],
					images: [oneCell(1, 1, a), oneCell(2, 2, aAgain), oneCell(3, 3, b)],
				},
			],
		});
		const media = entryNames(bytes).filter((n) => n.startsWith("xl/media/"));
		expect(media).toEqual(["xl/media/image1.png", "xl/media/image2.png"]); // 3 pics, 2 parts
	});

	// Review regression: identical bytes declared as two different image types are two different media
	// parts — dedup must NOT merge them, or the second picture's rel points at a file never written.
	it("does not merge identical bytes across different mime types", async () => {
		const raw = new Uint8Array([1, 2, 3, 4]);
		const input: WorkbookInput = {
			sheets: [
				{
					name: "S1",
					rows: [[1]],
					images: [
						{
							anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
							bytes: raw,
							mime: "image/png",
						},
					],
				},
				{
					name: "S2",
					rows: [[1]],
					images: [
						{
							anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
							bytes: raw,
							mime: "image/jpeg",
						},
					],
				},
			],
		};
		const bytes = await writeXlsx(input);
		const media = entryNames(bytes).filter((n) => n.startsWith("xl/media/"));
		expect(media).toEqual(["xl/media/image1.png", "xl/media/image2.jpeg"]); // two parts, not one
		const wb = await openXlsx(bytes);
		expect((await wb.sheet("S1").images()).length).toBe(1);
		expect((await wb.sheet("S2").images()).length).toBe(1); // was silently lost before the fix
	});

	it("streamed and buffered writers read back to the same images", async () => {
		const images = [
			oneCell(1, 1, PNG(1), { name: "A" }),
			{
				anchor: { from: { col: 3, row: 3 }, to: { col: 5, row: 5 } },
				bytes: new Uint8Array([0xff, 0xd8, 9]),
				mime: "image/jpeg",
			} satisfies SheetImage,
		];
		const buffered = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[1]], images }] }),
		);
		const streamed = await openXlsx(
			await drain(streamXlsx({ sheets: [{ name: "S", rows: [[1]], images }] })),
		);
		expect(await streamed.sheet("S").images()).toEqual(await buffered.sheet("S").images());
	});

	it("emits no drawing or media parts when a sheet has no images (unused emits nothing)", async () => {
		const withImages = entryNames(
			await writeXlsx({
				sheets: [{ name: "S", rows: [[1]], images: [oneCell(1, 1, PNG(1))] }],
			}),
		);
		const without = entryNames(await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] }));
		expect(withImages.some((n) => n.includes("/drawings/") || n.includes("/media/"))).toBe(
			true,
		);
		expect(without.some((n) => n.includes("/drawings/") || n.includes("/media/"))).toBe(false);
		expect(without.some((n) => n.startsWith("xl/comments"))).toBe(false);
	});

	it("reads the bytes reference exactly once (TOCTOU — a flip-getter can't swap the buffer)", async () => {
		const first = PNG(7);
		const second = PNG(8);
		let reads = 0;
		const image = {
			anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
			mime: "image/png",
			get bytes() {
				reads++;
				return reads === 1 ? first : second;
			},
		};
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[1]], images: [image] }] }),
		);
		const [img] = await wb.sheet("S").images();
		expect(img?.bytes).toEqual(first); // the buffer validated at read #1, not the flipped one
	});

	describe("validation rejects malformed images (typed, naming the sheet + index)", () => {
		const write = (images: unknown[]): Promise<Uint8Array> =>
			writeXlsx({ sheets: [{ name: "Sh", rows: [[1]], images: images as SheetImage[] }] });

		it.each([
			[
				"a bad mime",
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "image/webp",
					},
				],
				/image\/png/,
			],
			// Review regression: a mime that is an Object.prototype key must not slip past the allowlist.
			[
				'the prototype key "constructor" as mime',
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "constructor",
					},
				],
				/image\/png/,
			],
			[
				'the prototype key "__proto__" as mime',
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "__proto__",
					},
				],
				/image\/png/,
			],
			[
				"empty bytes",
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: new Uint8Array(0),
						mime: "image/png",
					},
				],
				/non-empty Uint8Array/,
			],
			[
				"non-Uint8Array bytes",
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: [1, 2, 3],
						mime: "image/png",
					},
				],
				/non-empty Uint8Array/,
			],
			[
				"both to and ext",
				[
					{
						anchor: {
							from: { col: 1, row: 1 },
							to: { col: 2, row: 2 },
							ext: { cx: 1, cy: 1 },
						},
						bytes: PNG(1),
						mime: "image/png",
					},
				],
				/exactly one of/,
			],
			[
				"neither to nor ext",
				[{ anchor: { from: { col: 1, row: 1 } }, bytes: PNG(1), mime: "image/png" }],
				/exactly one of/,
			],
			[
				"col out of grid",
				[
					{
						anchor: { from: { col: 0, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "image/png",
					},
				],
				/col must be an integer column/,
			],
			[
				"a negative EMU",
				[
					{
						anchor: { from: { col: 1, row: 1, colOff: -1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "image/png",
					},
				],
				/EMU/,
			],
			[
				"an over-range EMU",
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 0x80000000, cy: 1 } },
						bytes: PNG(1),
						mime: "image/png",
					},
				],
				/EMU/,
			],
			[
				"an unknown property",
				[
					{
						anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
						bytes: PNG(1),
						mime: "image/png",
						nope: 1,
					},
				],
				/unknown property/,
			],
			[
				"a bad editAs",
				[
					{
						anchor: { from: { col: 1, row: 1 }, to: { col: 2, row: 2 }, editAs: "wat" },
						bytes: PNG(1),
						mime: "image/png",
					},
				],
				/editAs/,
			],
			["a non-object image", [42], /must be an object/],
		])("rejects %s", async (_label, images, pattern) => {
			await expect(write(images)).rejects.toThrow(pattern as RegExp);
		});
	});
});
