import { openXlsx, streamSheetRows, XlsxError } from "@openjsxl/core";
import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { assertSameRows, replayAll, rowsOrTyped, streamMatchesBuffered } from "../replay";

// F11.47: the replay runs the ZIP openers with checksums on too, and compares streamSheetRows with the
// buffered rows of every sheet, plain and checked.

const LANES = [
	"openXlsx",
	"openXlsb",
	"openOds",
	"openCsv",
	"detect",
	"openXlsx+crc",
	"openXlsb+crc",
	"openOds+crc",
	"streamSheetRows",
	"streamSheetRows+crc",
];

// A copy of `bytes` whose central-directory CRC-32 for every entry under `prefix` is wrong.
function withBadCrc(bytes: Uint8Array, prefix: string): Uint8Array {
	const out = bytes.slice();
	const view = new DataView(out.buffer);
	let eocd = out.length - 22;
	while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd--;
	let pos = view.getUint32(eocd + 16, true);
	for (let i = view.getUint16(eocd + 10, true); i > 0; i--) {
		const nameLength = view.getUint16(pos + 28, true);
		const name = new TextDecoder().decode(out.subarray(pos + 46, pos + 46 + nameLength));
		if (name.startsWith(prefix))
			view.setUint32(pos + 16, view.getUint32(pos + 16, true) ^ 1, true);
		pos += 46 + nameLength + view.getUint16(pos + 30, true) + view.getUint16(pos + 32, true);
	}
	return out;
}

const verdicts = async (bytes: Uint8Array) =>
	Object.fromEntries((await replayAll(bytes)).outcomes.map((o) => [o.opener, o.verdict]));

describe("replayAll lanes (F11.47)", () => {
	it("runs every lane, and an intact xlsx passes the checked and streaming lanes", async () => {
		const result = await replayAll(await loadFixture("basic.xlsx"));
		expect(result.outcomes.map((o) => o.opener)).toEqual(LANES);
		expect(result.crashed).toBe(false);
		const v = await verdicts(await loadFixture("basic.xlsx"));
		expect([
			v.openXlsx,
			v["openXlsx+crc"],
			v.streamSheetRows,
			v["streamSheetRows+crc"],
		]).toEqual(["ok", "ok", "ok", "ok"]);
	});

	it("fails only the checked lanes, typed, when a worksheet's CRC-32 is wrong", async () => {
		const damaged = withBadCrc(await loadFixture("basic.xlsx"), "xl/worksheets/");
		const v = await verdicts(damaged);
		expect([v.openXlsx, v.streamSheetRows]).toEqual(["ok", "ok"]);
		expect([v["openXlsx+crc"], v["streamSheetRows+crc"]]).toEqual(["typed", "typed"]);
		// The stream's own check, at the end of the entry, refuses the same sheet.
		expect(await rowsOrTyped(streamSheetRows(damaged))).not.toBe("typed");
		expect(
			await rowsOrTyped(streamSheetRows(damaged, undefined, { verifyChecksums: true })),
		).toBe("typed");
	});

	it("checks the xlsb and ods openers with checksums on too", async () => {
		for (const [fixture, lane] of [
			["xlsb-basic.xlsb", "openXlsb"],
			["odf-basic.ods", "openOds"],
		] as const) {
			const intact = await verdicts(await loadFixture(fixture));
			expect([intact[lane], intact[`${lane}+crc`]]).toEqual(["ok", "ok"]);
			const damaged = await verdicts(withBadCrc(await loadFixture(fixture), ""));
			expect([damaged[lane], damaged[`${lane}+crc`]]).toEqual(["ok", "typed"]);
		}
	});
});

describe("rowsOrTyped", () => {
	it("is the rows as JSON text, or typed for an XlsxError", async () => {
		const book = await openXlsx(await loadFixture("basic.xlsx"));
		const name = book.sheets[0]?.name ?? "";
		const text = await rowsOrTyped(book.sheet(name).rows());
		const rows: unknown[] = [];
		for await (const row of book.sheet(name).rows()) rows.push(row);
		expect(text).toBe(JSON.stringify(rows));
		const failing = async function* () {
			yield* [];
			throw new XlsxError("corrupt-zip", "cut");
		};
		expect(await rowsOrTyped(failing())).toBe("typed");
	});

	it("rethrows any other failure, so the lane reports a crash", async () => {
		const crashing = async function* () {
			yield* [];
			throw new TypeError("boom");
		};
		await expect(rowsOrTyped(crashing())).rejects.toThrow(TypeError);
	});
});

describe("assertSameRows", () => {
	it("returns for equal reads and throws a plain Error naming the sheet otherwise", () => {
		expect(() => assertSameRows("S", "typed", "typed")).not.toThrow();
		expect(() => assertSameRows("S", "[1]", "[1]")).not.toThrow();
		expect(() => assertSameRows("S", "[]", "typed")).toThrow(
			'streamSheetRows and openXlsx disagree on sheet "S": typed vs []',
		);
		expect(() => assertSameRows("S", "[1]", "[2]")).toThrow(
			expect.objectContaining({ name: "Error" }),
		);
	});
});

describe("streamMatchesBuffered", () => {
	it("resolves when the stream gives every sheet's buffered rows, plain and checked", async () => {
		for (const fixture of ["basic.xlsx", "equiv.xlsx", "any_sheets.xlsx"]) {
			const bytes = await loadFixture(fixture);
			await expect(streamMatchesBuffered(bytes, {})).resolves.toBeUndefined();
			await expect(
				streamMatchesBuffered(bytes, { verifyChecksums: true }),
			).resolves.toBeUndefined();
			const book = await openXlsx(bytes);
			for (const info of book.sheets) {
				expect(await rowsOrTyped(streamSheetRows(bytes, info.name))).toBe(
					await rowsOrTyped(book.sheet(info.name).rows()),
				);
			}
		}
	});

	it("streams every sheet the buffered open loaded", async () => {
		// Each open reads its options once (F11.41(a)): one buffered open, then one stream per sheet.
		const bytes = await loadFixture("any_sheets.xlsx");
		const sheets = (await openXlsx(bytes)).sheets.length;
		expect(sheets).toBeGreaterThan(1);
		let reads = 0;
		const options = {
			get maxPartBytes() {
				reads++;
				return 2 ** 31;
			},
		};
		await streamMatchesBuffered(bytes, options);
		expect(reads).toBe(1 + sheets);
	});

	it("throws a plain Error when the stream disagrees with the buffered rows", async () => {
		// The buffered open reads maxPartBytes 2^31; every stream after it reads 1 and fails typed.
		let reads = 0;
		const options = {
			get maxPartBytes() {
				return reads++ === 0 ? 2 ** 31 : 1;
			},
		};
		const result = streamMatchesBuffered(await loadFixture("basic.xlsx"), options);
		await expect(result).rejects.toThrow(/streamSheetRows and openXlsx disagree on sheet/);
		await expect(result).rejects.toThrow(expect.objectContaining({ name: "Error" }));
	});

	it("still streams the first sheet when the buffered open fails, then rethrows its error", async () => {
		const damaged = withBadCrc(await loadFixture("basic.xlsx"), "xl/worksheets/");
		let reads = 0;
		const options = {
			verifyChecksums: true,
			get maxPartBytes() {
				reads++;
				return 2 ** 31;
			},
		};
		await expect(streamMatchesBuffered(damaged, options)).rejects.toMatchObject({
			name: "XlsxError",
			code: "corrupt-zip",
		});
		expect(reads).toBe(2); // the buffered open, then the stream of the first sheet
	});
});
