import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { detectSpreadsheetFormat } from "../detect";
import { openOds } from "../ods";
import { openXlsx, type ReadOptions, streamSheetRows, type Workbook } from "../workbook";
import { openXlsb } from "../xlsb";

// F11.41(a), owner call D17: every reader that takes ReadOptions, and detectSpreadsheetFormat, checks
// them once before it reads the file. A limit is a number, not NaN, at least 0, with Infinity for no
// limit; verifyChecksums is a boolean; undefined and null keep the default; anything else fails typed
// invalid-input, on any input, so no option value can switch a guard off.

const typed = (code: string) => expect.objectContaining({ name: "XlsxError", code });

const LIMITS = ["maxPartBytes", "maxCompressionRatio", "maxTotalBytes"] as const;
const BAD_LIMITS: readonly unknown[] = [Number.NaN, "abc", "100", -1, Number.NEGATIVE_INFINITY, {}];
const BAD_FLAGS: readonly unknown[] = ["true", "false", 1, 0];
const NOT_A_ZIP = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);

// Every sheet's name and cell values, in tab order.
async function snapshot(book: Workbook): Promise<unknown[]> {
	const out: unknown[] = [];
	for (const info of book.sheets) {
		const rows: unknown[] = [];
		for await (const row of book.sheet(info.name).rows())
			rows.push(row.cells.map((c) => c.value));
		out.push([info.name, rows]);
	}
	return out;
}

async function streamed(bytes: Uint8Array, options?: ReadOptions): Promise<unknown[]> {
	const rows: unknown[] = [];
	for await (const row of streamSheetRows(bytes, undefined, options))
		rows.push(row.cells.map((c) => c.value));
	return rows;
}

type Read = (bytes: Uint8Array, options?: ReadOptions) => Promise<unknown>;

const READERS: readonly (readonly [string, string, Read])[] = [
	["openXlsx", "equiv.xlsx", async (bytes, options) => snapshot(await openXlsx(bytes, options))],
	["openXlsb", "equiv.xlsb", async (bytes, options) => snapshot(await openXlsb(bytes, options))],
	["openOds", "equiv.ods", async (bytes, options) => snapshot(await openOds(bytes, options))],
	["streamSheetRows", "equiv.xlsx", streamed],
];

const DETECTED: readonly (readonly [string, string])[] = [
	["equiv.xlsx", "xlsx"],
	["equiv.xlsb", "xlsb"],
	["equiv.ods", "ods"],
	["equiv.csv", "csv"],
];

describe.each(READERS)(
	"%s checks ReadOptions before it reads (F11.41(a))",
	(_name, fixture, read) => {
		it.each(LIMITS)(
			"refuses a bad %s, typed, on its file and on bytes that are not a ZIP",
			async (key) => {
				const bytes = await loadFixture(fixture);
				for (const value of BAD_LIMITS) {
					const options = { [key]: value } as never;
					await expect(read(bytes, options)).rejects.toThrow(typed("invalid-input"));
					await expect(read(NOT_A_ZIP, options)).rejects.toThrow(typed("invalid-input"));
				}
			},
		);

		it("refuses a verifyChecksums that is not a boolean, typed", async () => {
			const bytes = await loadFixture(fixture);
			for (const value of BAD_FLAGS) {
				const options = { verifyChecksums: value } as never;
				await expect(read(bytes, options)).rejects.toThrow(typed("invalid-input"));
				await expect(read(NOT_A_ZIP, options)).rejects.toThrow(typed("invalid-input"));
			}
		});

		it.each(LIMITS)(
			"reads as with no options when %s is Infinity, undefined or null",
			async (key) => {
				const bytes = await loadFixture(fixture);
				const plain = await read(bytes);
				expect(plain).not.toEqual([]);
				for (const value of [Number.POSITIVE_INFINITY, undefined, null]) {
					expect(await read(bytes, { [key]: value } as never)).toEqual(plain);
				}
			},
		);

		it("reads as with no options when verifyChecksums is true, false, undefined or null", async () => {
			const bytes = await loadFixture(fixture);
			const plain = await read(bytes);
			for (const value of [true, false, undefined, null]) {
				expect(await read(bytes, { verifyChecksums: value } as never)).toEqual(plain);
			}
		});

		it("takes 0 as a limit: no part fits 0 bytes, and a 0 ratio keeps its 8 MiB floor", async () => {
			const bytes = await loadFixture(fixture);
			await expect(read(bytes, { maxPartBytes: 0 })).rejects.toThrow(typed("part-too-large"));
			await expect(read(bytes, { maxTotalBytes: 0 })).rejects.toThrow(
				typed("part-too-large"),
			);
			expect(await read(bytes, { maxCompressionRatio: 0 })).toEqual(await read(bytes));
		});

		it("reads each option once per call", async () => {
			const bytes = await loadFixture(fixture);
			const reads = {
				maxPartBytes: 0,
				maxCompressionRatio: 0,
				maxTotalBytes: 0,
				verifyChecksums: 0,
			};
			// A second read of any option would see NaN, or a non-boolean, and fail.
			const options = {
				get maxPartBytes() {
					return reads.maxPartBytes++ === 0 ? 2 ** 31 : Number.NaN;
				},
				get maxCompressionRatio() {
					return reads.maxCompressionRatio++ === 0 ? 300 : Number.NaN;
				},
				get maxTotalBytes() {
					return reads.maxTotalBytes++ === 0 ? Number.POSITIVE_INFINITY : Number.NaN;
				},
				get verifyChecksums() {
					return reads.verifyChecksums++ === 0 ? true : "no";
				},
			};
			expect(await read(bytes, options as never)).toEqual(await read(bytes));
			expect(reads).toEqual({
				maxPartBytes: 1,
				maxCompressionRatio: 1,
				maxTotalBytes: 1,
				verifyChecksums: 1,
			});
		});
	},
);

describe("detectSpreadsheetFormat checks ReadOptions before it looks at the bytes (F11.41(a))", () => {
	it.each(DETECTED)(
		"refuses a bad option, typed, for %s, rather than returning a format or undefined",
		async (fixture) => {
			const bytes = await loadFixture(fixture);
			for (const key of LIMITS) {
				for (const value of BAD_LIMITS) {
					await expect(
						detectSpreadsheetFormat(bytes, { [key]: value } as never),
					).rejects.toThrow(typed("invalid-input"));
				}
			}
			for (const value of BAD_FLAGS) {
				await expect(
					detectSpreadsheetFormat(bytes, { verifyChecksums: value } as never),
				).rejects.toThrow(typed("invalid-input"));
			}
		},
	);

	it("refuses a bad option for empty input and binary noise too", async () => {
		for (const bytes of [new Uint8Array(0), new Uint8Array([0, 1, 2, 3])]) {
			await expect(
				detectSpreadsheetFormat(bytes, { maxPartBytes: Number.NaN }),
			).rejects.toThrow(typed("invalid-input"));
			expect(await detectSpreadsheetFormat(bytes)).toBeUndefined();
		}
	});

	it.each(DETECTED)(
		"detects %s as with no options for Infinity, undefined and null",
		async (fixture, format) => {
			const bytes = await loadFixture(fixture);
			expect(await detectSpreadsheetFormat(bytes)).toBe(format);
			for (const key of LIMITS) {
				for (const value of [Number.POSITIVE_INFINITY, undefined, null]) {
					expect(await detectSpreadsheetFormat(bytes, { [key]: value } as never)).toBe(
						format,
					);
				}
			}
			for (const value of [true, false, undefined, null]) {
				expect(
					await detectSpreadsheetFormat(bytes, { verifyChecksums: value } as never),
				).toBe(format);
			}
		},
	);

	it.each(DETECTED)("reads each option once for %s", async (fixture, format) => {
		const bytes = await loadFixture(fixture);
		let reads = 0;
		const options = {
			get maxPartBytes() {
				return reads++ === 0 ? 2 ** 31 : Number.NaN;
			},
		};
		expect(await detectSpreadsheetFormat(bytes, options)).toBe(format);
		expect(reads).toBe(1);
	});
});
