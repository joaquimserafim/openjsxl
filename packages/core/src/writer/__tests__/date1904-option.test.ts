import { describe, expect, it } from "vitest";
import { openZip } from "../../zip";
import { optionsDate1904 } from "../parts";
import { streamXlsx } from "../stream";
import type { WorkbookInput, WriteOptions } from "../types";
import { writeXlsx } from "../workbook";

// F11.99: F11.10 passed the caller's `WriteOptions.date1904` through as it was, and both writers test
// it for truthiness, so a JavaScript caller's "false", "0", 1, "yes" or "true" selected the 1904
// system and shifted every date by 1,462 days. 1.0.0 read `options?.date1904 === true`. Only `true`
// is 1904 again. `null` and `undefined` defer to WorkbookInput.date1904, and any other value
// overrides it, as `false` does.

// Every flag value but `true`, past the declared boolean type: each must be the 1900 system.
const NOT_TRUE: readonly unknown[] = [
	false,
	"false",
	"0",
	"true",
	"yes",
	1,
	0,
	"",
	Number.NaN,
	{},
	[],
	new Boolean(true),
];

// An options object whose own `date1904` holds `value`.
const withFlag = (value: unknown): WriteOptions =>
	Object.defineProperty({}, "date1904", {
		value,
		enumerable: true,
		writable: true,
		configurable: true,
	});

// An options object whose own `date1904` getter returns `first`, then `then` on every later read.
function flagGetter(
	first: unknown,
	then: unknown,
): { readonly options: WriteOptions; readonly reads: () => number } {
	let reads = 0;
	const options = Object.defineProperty({}, "date1904", {
		enumerable: true,
		configurable: true,
		get: () => (reads++ === 0 ? first : then),
	});
	return { options, reads: () => reads };
}

async function streamed(input: WorkbookInput, options?: WriteOptions): Promise<Uint8Array> {
	const chunks: Uint8Array[] = [];
	const reader = streamXlsx(input, options).getReader();
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
	}
	const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.length;
	}
	return bytes;
}

const writers = [
	{
		name: "writeXlsx",
		write: (input: WorkbookInput, options?: WriteOptions) => writeXlsx(input, options),
	},
	{ name: "streamXlsx", write: streamed },
];

const decoder = new TextDecoder();
const workbookXml = async (bytes: Uint8Array) =>
	decoder.decode(await openZip(bytes).read("xl/workbook.xml"));

// A1 is a date, so its serial differs by 1,462 between the systems; B1 caches a raw serial.
const rows = [[new Date(Date.UTC(1904, 0, 1)), { formula: "A1+30", value: 30 }]];
const plain = (): WorkbookInput => ({ sheets: [{ name: "S", rows }] });
const flagged = (): WorkbookInput => ({ sheets: [{ name: "S", rows }], date1904: true });

describe("optionsDate1904 — only true is the 1904 system (F11.99)", () => {
	it("returns true for true", () => {
		expect(optionsDate1904({ date1904: true })).toBe(true);
		expect(optionsDate1904(withFlag(true))).toBe(true);
	});

	it("returns false for every other value", () => {
		for (const value of NOT_TRUE) {
			expect(optionsDate1904(withFlag(value)), String(value)).toBe(false);
		}
	});

	it("is undefined for a null or undefined flag, or none", () => {
		expect(optionsDate1904(withFlag(null))).toBeUndefined();
		expect(optionsDate1904(withFlag(undefined))).toBeUndefined();
		expect(optionsDate1904({})).toBeUndefined();
	});

	it("reads the flag once, and that read decides", () => {
		const { options, reads } = flagGetter("yes", true);
		expect(optionsDate1904(options)).toBe(false);
		expect(reads()).toBe(1);
	});
});

for (const { name, write } of writers) {
	describe(`${name} — options.date1904 (F11.99)`, () => {
		it("writes the bytes of a write without the option for every value but true", async () => {
			const baseline = await write(plain());
			for (const value of [...NOT_TRUE, null, undefined]) {
				const bytes = await write(plain(), withFlag(value));
				expect(Buffer.compare(bytes, baseline), String(value)).toBe(0);
			}
			for (const options of [{}, undefined]) {
				expect(Buffer.compare(await write(plain(), options), baseline)).toBe(0);
			}
			expect(await workbookXml(baseline)).not.toContain("date1904");
		});

		it("selects the 1904 system for true", async () => {
			const bytes = await write(plain(), withFlag(true));
			expect(await workbookXml(bytes)).toContain('<workbookPr date1904="1"/>');
			expect(Buffer.compare(bytes, await write(flagged()))).toBe(0);
		});

		it("defers to workbook.date1904 for null and undefined, and overrides it otherwise", async () => {
			const system1904 = await write(flagged());
			const system1900 = await write(plain());
			for (const value of [null, undefined]) {
				const bytes = await write(flagged(), withFlag(value));
				expect(Buffer.compare(bytes, system1904), String(value)).toBe(0);
			}
			for (const value of NOT_TRUE) {
				const bytes = await write(flagged(), withFlag(value));
				expect(Buffer.compare(bytes, system1900), String(value)).toBe(0);
			}
			const kept = await write(flagged(), withFlag(true));
			expect(Buffer.compare(kept, system1904)).toBe(0);
		});

		it("reads a getter flag once per write, and its first value decides", async () => {
			const system1900 = await write(plain());
			for (const input of [plain(), flagged()]) {
				const { options, reads } = flagGetter("false", true);
				const bytes = await write(input, options);
				expect(Buffer.compare(bytes, system1900)).toBe(0);
				expect(reads()).toBe(1);
			}
		});
	});
}
