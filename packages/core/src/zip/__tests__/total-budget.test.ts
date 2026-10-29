import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import { budgetMeter, localEntrySpan, openZip, totalBudget } from "../central-directory";
import { centralRecord } from "./entry-overlap-fixtures";

// F11.53: `maxTotalBytes` bounds the decompressed bytes read through one archive. Each read is charged
// the most it can produce before it starts, stored and inflated parts alike, and the read that would
// pass the budget fails typed. The option is validated on its own with F11.41(a)'s rule.

const typed = (code: string) => expect.objectContaining({ name: "XlsxError", code });

// Three parts: two that deflate well and one of noise that stays stored.
async function archive() {
	let seed = 7;
	const noise = Uint8Array.from({ length: 1000 }, () => {
		seed = (seed * 48271) % 2147483647;
		return seed & 0xff;
	});
	const text = (n: number) => new TextEncoder().encode("<row/>".repeat(n));
	const parts = [
		{ name: "a.xml", data: text(2000) },
		{ name: "b.bin", data: noise },
		{ name: "c.xml", data: text(1000) },
	];
	const bytes = await writeZip(parts);
	const sizes = parts.map((part) => part.data.byteLength);
	return { bytes, sizes, total: sizes.reduce((x, y) => x + y, 0) };
}

async function drain(stream: AsyncGenerator<Uint8Array>) {
	let n = 0;
	for await (const chunk of stream) n += chunk.byteLength;
	return n;
}

describe("totalBudget", () => {
	it("is unlimited when absent, undefined or null", () => {
		expect(totalBudget(undefined)).toBe(Number.POSITIVE_INFINITY);
		expect(totalBudget(null)).toBe(Number.POSITIVE_INFINITY);
	});

	it("takes a number of bytes, at least 0, Infinity included", () => {
		expect(totalBudget(0)).toBe(0);
		expect(totalBudget(64 * 1024 * 1024)).toBe(67108864);
		expect(totalBudget(Number.POSITIVE_INFINITY)).toBe(Number.POSITIVE_INFINITY);
	});

	it("refuses anything else, typed", () => {
		for (const value of [Number.NaN, -1, Number.NEGATIVE_INFINITY, "100", true, {}, 10n]) {
			expect(() => totalBudget(value)).toThrow(typed("invalid-input"));
		}
	});
});

describe("budgetMeter", () => {
	it("allows reads up to the limit exactly and refuses the one past it", () => {
		const charge = budgetMeter(100);
		charge("a.xml", 60);
		charge("b.xml", 40);
		expect(() => charge("c.xml", 1)).toThrow(typed("part-too-large"));
		expect(() => charge("c.xml", 1)).toThrow(/c\.xml.*101 bytes.*100-byte maxTotalBytes/);
	});

	it("still allows an empty read once the budget is spent, and never counts a refused one", () => {
		const charge = budgetMeter(10);
		charge("a.xml", 10);
		charge("b.xml", 0);
		expect(() => charge("c.xml", 5)).toThrow(typed("part-too-large"));
		charge("d.xml", 0);
	});

	it("never refuses without a limit", () => {
		const charge = budgetMeter(Number.POSITIVE_INFINITY);
		for (let i = 0; i < 3; i++) charge("a.xml", Number.MAX_SAFE_INTEGER);
		charge("b.xml", Number.NaN);
	});

	it("refuses a bound that is not a number under a finite limit", () => {
		expect(() => budgetMeter(100)("a.xml", Number.NaN)).toThrow(typed("part-too-large"));
		expect(() => budgetMeter(0)("a.xml", Number.NaN)).toThrow(/a\.xml has no size bound/);
	});
});

describe("openZip's maxTotalBytes", () => {
	it("leaves every read as it was when the parts fit, or with no budget", async () => {
		const { bytes, sizes, total } = await archive();
		for (const options of [
			{},
			{ maxTotalBytes: total },
			{ maxTotalBytes: undefined as never },
		]) {
			const zip = openZip(bytes, options);
			const read = await Promise.all(
				["a.xml", "b.bin", "c.xml"].map((name) => zip.read(name)),
			);
			expect(read.map((part) => part.byteLength)).toEqual(sizes);
		}
	});

	it("refuses the read that passes the budget across several parts, buffered or streamed", async () => {
		const { bytes, sizes, total } = await archive();
		const buffered = openZip(bytes, { maxTotalBytes: total - 1 });
		expect((await buffered.read("a.xml")).byteLength).toBe(sizes[0]);
		expect((await buffered.read("b.bin")).byteLength).toBe(sizes[1]);
		await expect(buffered.read("c.xml")).rejects.toThrow(typed("part-too-large"));

		const streamed = openZip(bytes, { maxTotalBytes: total - 1 });
		expect(await drain(streamed.readStream("a.xml"))).toBe(sizes[0]);
		expect(await drain(streamed.readStream("b.bin"))).toBe(sizes[1]);
		await expect(drain(streamed.readStream("c.xml"))).rejects.toThrow(typed("part-too-large"));

		const mixed = openZip(bytes, { maxTotalBytes: total - 1 });
		await mixed.read("a.xml");
		await drain(mixed.readStream("b.bin"));
		await expect(mixed.read("c.xml")).rejects.toThrow(typed("part-too-large"));
	});

	it("counts a stored part at the bytes it holds, whatever size it declares", async () => {
		const { bytes, sizes } = await archive();
		// b.bin stores 1000 bytes but now declares 5.
		const lying = bytes.slice();
		new DataView(lying.buffer).setUint32(centralRecord(bytes, "b.bin").offset + 24, 5, true);
		const zip = openZip(lying, { maxTotalBytes: (sizes[1] ?? 0) - 1 });
		let delivered = 0;
		await expect(
			(async () => {
				for await (const chunk of zip.readStream("b.bin")) delivered += chunk.byteLength;
			})(),
		).rejects.toThrow(typed("part-too-large"));
		expect(delivered).toBe(0);
	});

	it("counts a stored part at its own size", async () => {
		const { bytes, sizes } = await archive();
		const zip = openZip(bytes, { maxTotalBytes: (sizes[1] ?? 0) - 1 });
		expect(zip.entries.get("b.bin")?.method).toBe(0);
		await expect(zip.read("b.bin")).rejects.toThrow(typed("part-too-large"));
		await expect(drain(zip.readStream("b.bin"))).rejects.toThrow(typed("part-too-large"));
	});

	it("counts a part read twice twice, and a stream cut short at its whole size", async () => {
		const { bytes, sizes } = await archive();
		const a = sizes[0] ?? 0;
		const twice = openZip(bytes, { maxTotalBytes: 2 * a - 1 });
		await twice.read("a.xml");
		await expect(twice.read("a.xml")).rejects.toThrow(typed("part-too-large"));

		// A part large enough to inflate in many chunks, cut after the first: the whole part is still
		// counted, so a small part that would fit beside its first chunk alone is refused.
		const big = new TextEncoder().encode("<row/>".repeat(200_000));
		const small = new TextEncoder().encode("<row/>".repeat(200));
		const many = await writeZip([
			{ name: "big.xml", data: big },
			{ name: "small.xml", data: small },
		]);
		const cut = openZip(many, { maxTotalBytes: big.byteLength + small.byteLength - 1 });
		let chunks = 0;
		for await (const chunk of cut.readStream("big.xml")) {
			chunks++;
			expect(chunk.byteLength).toBeLessThan(big.byteLength);
			break;
		}
		expect(chunks).toBe(1);
		await expect(cut.read("small.xml")).rejects.toThrow(typed("part-too-large"));
	});

	it("refuses before inflating, so a corrupt part over the budget fails on the budget", async () => {
		const { bytes } = await archive();
		const zip = openZip(bytes);
		const entry = zip.entries.get("a.xml");
		const span =
			entry === undefined ? undefined : localEntrySpan(new DataView(bytes.buffer), entry);
		expect(entry?.method).toBe(8);
		const corrupt = bytes.slice();
		corrupt.fill(0xff, span?.dataStart ?? 0, (span?.dataStart ?? 0) + 8);
		await expect(openZip(corrupt).read("a.xml")).rejects.toThrow(typed("corrupt-zip"));
		await expect(openZip(corrupt, { maxTotalBytes: 100 }).read("a.xml")).rejects.toThrow(
			typed("part-too-large"),
		);
	});

	it("refuses an inflated part under a finite budget when a sibling option is NaN", async () => {
		const { bytes, sizes } = await archive();
		for (const sibling of [{ maxPartBytes: Number.NaN }, { maxCompressionRatio: Number.NaN }]) {
			const zip = openZip(bytes, { ...sibling, maxTotalBytes: 1_000_000 });
			await expect(zip.read("a.xml")).rejects.toThrow(typed("part-too-large"));
			await expect(drain(zip.readStream("c.xml"))).rejects.toThrow(typed("part-too-large"));
			expect((await zip.read("b.bin")).byteLength).toBe(sizes[1]);
			// Without a budget, such a read is as before.
			expect((await openZip(bytes, sibling).read("a.xml")).byteLength).toBe(sizes[0]);
		}
	});

	it("validates the option before anything else, even on bytes that are not a ZIP", () => {
		expect(() => openZip(new Uint8Array(4), { maxTotalBytes: Number.NaN })).toThrow(
			typed("invalid-input"),
		);
		expect(() => openZip(new Uint8Array(4), { maxTotalBytes: null as never })).toThrow(
			typed("not-a-zip"),
		);
	});

	it("reads the option once", async () => {
		const { bytes, total } = await archive();
		let reads = 0;
		const options = {
			get maxTotalBytes() {
				reads++;
				return reads === 1 ? total : 0;
			},
		};
		const zip = openZip(bytes, options);
		await zip.read("a.xml");
		await zip.read("c.xml");
		expect(reads).toBe(1);
	});
});
