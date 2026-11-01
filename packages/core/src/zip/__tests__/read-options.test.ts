import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import {
	booleanOption,
	DEFAULT_MAX_COMPRESSION_RATIO,
	DEFAULT_MAX_PART_BYTES,
	inflateCap,
	limitOption,
	openZip,
	RATIO_FLOOR_BYTES,
	zipLimits,
} from "../central-directory";
import { centralRecord } from "./entry-overlap-fixtures";

// F11.41(a), owner call D17: every ZIP read option is read once and checked before the archive is
// parsed. A limit is a number, not NaN, at least 0, with Infinity for no limit; verifyChecksums is a
// boolean; undefined and null keep the default; anything else fails typed invalid-input. The
// declared size caps an inflate whatever the options are.

const typed = (code: string) => expect.objectContaining({ name: "XlsxError", code });

const DEFAULTS = {
	maxPartBytes: DEFAULT_MAX_PART_BYTES,
	maxCompressionRatio: DEFAULT_MAX_COMPRESSION_RATIO,
	maxTotalBytes: Number.POSITIVE_INFINITY,
	verifyChecksums: false,
};
const LIMITS = ["maxPartBytes", "maxCompressionRatio", "maxTotalBytes"] as const;
const BAD_LIMITS: readonly unknown[] = [
	Number.NaN,
	"abc",
	"100",
	"",
	-1,
	-0.5,
	Number.NEGATIVE_INFINITY,
	true,
	false,
	{},
	[],
	10n,
	() => 1,
];
const BAD_FLAGS: readonly unknown[] = ["true", "false", "", 1, 0, Number.NaN, {}, []];

async function drain(stream: AsyncGenerator<Uint8Array>): Promise<number> {
	let n = 0;
	for await (const chunk of stream) n += chunk.byteLength;
	return n;
}

// A one-part archive whose deflated part holds `size` zero bytes but declares `declared`.
async function lyingArchive(size: number, declared: number): Promise<Uint8Array> {
	const bytes = await writeZip([{ name: "a.xml", data: new Uint8Array(size) }]);
	const out = bytes.slice();
	new DataView(out.buffer).setUint32(centralRecord(bytes, "a.xml").offset + 24, declared, true);
	return out;
}

describe("zipLimits", () => {
	it("keeps every default for no options, an empty object, or undefined and null fields", () => {
		expect(zipLimits(undefined)).toEqual(DEFAULTS);
		expect(zipLimits(null as never)).toEqual(DEFAULTS);
		expect(zipLimits({})).toEqual(DEFAULTS);
		const absent = {
			maxPartBytes: undefined,
			maxCompressionRatio: undefined,
			maxTotalBytes: undefined,
			verifyChecksums: undefined,
		};
		expect(zipLimits(absent as never)).toEqual(DEFAULTS);
		const nulls = {
			maxPartBytes: null,
			maxCompressionRatio: null,
			maxTotalBytes: null,
			verifyChecksums: null,
		};
		expect(zipLimits(nulls as never)).toEqual(DEFAULTS);
	});

	it.each(LIMITS)("takes %s as any number at least 0, Infinity included", (key) => {
		for (const value of [0, -0, 1, 1.5, 2 ** 40, Number.MAX_VALUE, Number.POSITIVE_INFINITY]) {
			expect(zipLimits({ [key]: value })[key]).toBe(value);
		}
	});

	it.each(LIMITS)("refuses any other %s, typed, naming the option", (key) => {
		for (const value of BAD_LIMITS) {
			expect(() => zipLimits({ [key]: value as never })).toThrow(typed("invalid-input"));
			expect(() => zipLimits({ [key]: value as never })).toThrow(
				new RegExp(`^ReadOptions\\.${key} must be a number`),
			);
		}
	});

	it("takes verifyChecksums as true or false, and refuses anything else, typed", () => {
		expect(zipLimits({ verifyChecksums: true }).verifyChecksums).toBe(true);
		expect(zipLimits({ verifyChecksums: false }).verifyChecksums).toBe(false);
		for (const value of BAD_FLAGS) {
			expect(() => zipLimits({ verifyChecksums: value as never })).toThrow(
				typed("invalid-input"),
			);
			expect(() => zipLimits({ verifyChecksums: value as never })).toThrow(
				"ReadOptions.verifyChecksums must be true or false",
			);
		}
	});

	it("keeps F11.53's maxTotalBytes message", () => {
		expect(() => zipLimits({ maxTotalBytes: Number.NaN })).toThrow(
			"ReadOptions.maxTotalBytes must be a number of bytes, at least 0 (Infinity for no limit)",
		);
	});

	it("reads each option exactly once, so a getter cannot pass the check and then change", () => {
		const reads = {
			maxPartBytes: 0,
			maxCompressionRatio: 0,
			maxTotalBytes: 0,
			verifyChecksums: 0,
		};
		const options = {
			get maxPartBytes() {
				reads.maxPartBytes++;
				return reads.maxPartBytes === 1 ? 1000 : Number.NaN;
			},
			get maxCompressionRatio() {
				reads.maxCompressionRatio++;
				return reads.maxCompressionRatio === 1 ? 50 : Number.NaN;
			},
			get maxTotalBytes() {
				reads.maxTotalBytes++;
				return reads.maxTotalBytes === 1 ? 5000 : Number.NaN;
			},
			get verifyChecksums() {
				reads.verifyChecksums++;
				return reads.verifyChecksums === 1;
			},
		};
		expect(zipLimits(options)).toEqual({
			maxPartBytes: 1000,
			maxCompressionRatio: 50,
			maxTotalBytes: 5000,
			verifyChecksums: true,
		});
		expect(reads).toEqual({
			maxPartBytes: 1,
			maxCompressionRatio: 1,
			maxTotalBytes: 1,
			verifyChecksums: 1,
		});
	});

	it("returns a fresh object, so later changes to the caller's options do not reach it", () => {
		const options = { maxPartBytes: 10 };
		const limits = zipLimits(options);
		options.maxPartBytes = Number.NaN;
		expect(limits.maxPartBytes).toBe(10);
	});
});

describe("limitOption", () => {
	it("returns the fallback when absent and the value when it is a number at least 0", () => {
		expect(limitOption(undefined, 300, "m")).toBe(300);
		expect(limitOption(null, 300, "m")).toBe(300);
		expect(limitOption(0, 300, "m")).toBe(0);
		expect(limitOption(7.5, 300, "m")).toBe(7.5);
		expect(limitOption(Number.POSITIVE_INFINITY, 300, "m")).toBe(Number.POSITIVE_INFINITY);
	});

	it("throws invalid-input with the given message for anything else", () => {
		for (const value of BAD_LIMITS) {
			expect(() => limitOption(value, 300, "bad limit")).toThrow(typed("invalid-input"));
			expect(() => limitOption(value, 300, "bad limit")).toThrow("bad limit");
		}
	});
});

describe("booleanOption", () => {
	it("returns the fallback when absent and the value when it is a boolean", () => {
		expect(booleanOption(undefined, false, "m")).toBe(false);
		expect(booleanOption(null, true, "m")).toBe(true);
		expect(booleanOption(true, false, "m")).toBe(true);
		expect(booleanOption(false, true, "m")).toBe(false);
	});

	it("throws invalid-input with the given message for anything else", () => {
		for (const value of BAD_FLAGS) {
			expect(() => booleanOption(value, false, "bad flag")).toThrow(typed("invalid-input"));
			expect(() => booleanOption(value, false, "bad flag")).toThrow("bad flag");
		}
	});
});

describe("inflateCap", () => {
	it("is the smallest of the declared size, maxPartBytes and the ratio limit", () => {
		expect(inflateCap(100, 10, 2 ** 31, 300)).toBe(100);
		expect(inflateCap(100, 10, 50, 300)).toBe(50);
		expect(inflateCap(2 ** 30, 1000, 2 ** 31, 300)).toBe(RATIO_FLOOR_BYTES);
		expect(inflateCap(2 ** 30, 2 ** 20, 2 ** 31, 300)).toBe(314572800);
		expect(inflateCap(2 ** 30, 2 ** 20, 2 ** 31, 0)).toBe(RATIO_FLOOR_BYTES);
		expect(inflateCap(100, 10, 0, 300)).toBe(0);
	});

	it("is the declared size when both limits are Infinity", () => {
		const inf = Number.POSITIVE_INFINITY;
		expect(inflateCap(100, 10, inf, inf)).toBe(100);
		expect(inflateCap(2 ** 32 - 1, 2 ** 20, inf, inf)).toBe(2 ** 32 - 1);
	});

	it("keeps the declared size, and every other number, as a cap when a limit is not a number", () => {
		expect(inflateCap(100, 10, Number.NaN, Number.NaN)).toBe(100);
		expect(inflateCap(100, 10, Number.NaN, 300)).toBe(100);
		expect(inflateCap(2 ** 30, 1000, Number.NaN, 300)).toBe(RATIO_FLOOR_BYTES);
		expect(inflateCap(2 ** 30, 1000, 50, Number.NaN)).toBe(50);
		// 0 × Infinity is NaN: an Infinity ratio over an empty payload lifts nothing.
		expect(inflateCap(100, 0, 50, Number.POSITIVE_INFINITY)).toBe(50);
		expect(inflateCap(100, 0, Number.POSITIVE_INFINITY, Number.POSITIVE_INFINITY)).toBe(100);
	});
});

describe("openZip's read options (F11.41(a))", () => {
	it.each(LIMITS)("refuses a bad %s before parsing, even on bytes that are not a ZIP", (key) => {
		for (const value of [Number.NaN, "abc", -1]) {
			expect(() => openZip(new Uint8Array(4), { [key]: value as never })).toThrow(
				typed("invalid-input"),
			);
		}
		expect(() => openZip(new Uint8Array(4), { [key]: null as never })).toThrow(
			typed("not-a-zip"),
		);
	});

	it("refuses a verifyChecksums that is not a boolean before parsing", () => {
		expect(() => openZip(new Uint8Array(4), { verifyChecksums: "true" as never })).toThrow(
			typed("invalid-input"),
		);
		expect(() => openZip(new Uint8Array(4), { verifyChecksums: null as never })).toThrow(
			typed("not-a-zip"),
		);
	});

	it("stops a part that inflates past its declared size, even with both limits off", async () => {
		// 64 MiB of zeros deflates to about 64 KiB; the entry declares 100 bytes.
		const bytes = await lyingArchive(64 * 1024 * 1024, 100);
		const inf = Number.POSITIVE_INFINITY;
		for (const options of [
			undefined,
			{ maxPartBytes: inf, maxCompressionRatio: inf },
			{ maxPartBytes: null, maxCompressionRatio: null } as never,
		]) {
			const zip = openZip(bytes, options);
			await expect(zip.read("a.xml")).rejects.toThrow(typed("part-too-large"));
			await expect(zip.read("a.xml")).rejects.toThrow(/exceeds the 100-byte limit/);
			await expect(drain(zip.readStream("a.xml"))).rejects.toThrow(typed("part-too-large"));
		}
	});

	it("refuses NaN limits when the archive opens, instead of inflating the whole part", async () => {
		const bytes = await lyingArchive(64 * 1024 * 1024, 100);
		for (const options of [
			{ maxPartBytes: Number.NaN },
			{ maxCompressionRatio: Number.NaN },
			{ maxPartBytes: Number.NaN, maxCompressionRatio: Number.NaN },
		]) {
			expect(() => openZip(bytes, options)).toThrow(typed("invalid-input"));
		}
	});

	it("reads parts with the checked limits: 0 refuses every non-empty part, and a 0 ratio still has its floor", async () => {
		const data = new TextEncoder().encode("<row/>".repeat(1000));
		const bytes = await writeZip([{ name: "a.xml", data }]);
		await expect(openZip(bytes, { maxPartBytes: 0 }).read("a.xml")).rejects.toThrow(
			typed("part-too-large"),
		);
		expect(await openZip(bytes, { maxCompressionRatio: 0 }).read("a.xml")).toEqual(data);
		expect(await openZip(bytes, { verifyChecksums: true }).read("a.xml")).toEqual(data);
	});
});
