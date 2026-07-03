import { describe, expect, it } from "vitest";
import { openZip } from "../../zip";
import { type StreamPart, streamZip } from "../stream-zip";

// F5.1 — the streaming zip layer. Its data-descriptor archive must read back through openZip exactly
// like a buffered one; CRC correctness (which openZip trusts rather than verifies) is checked out of
// band by `unzip -t` and by openpyxl reading a streamed .xlsx.

const enc = new TextEncoder();
const dec = new TextDecoder();

async function collect(gen: AsyncGenerator<Uint8Array>): Promise<Uint8Array> {
	const chunks: Uint8Array[] = [];
	let total = 0;
	for await (const c of gen) {
		chunks.push(c);
		total += c.length;
	}
	const out = new Uint8Array(total);
	let o = 0;
	for (const c of chunks) {
		out.set(c, o);
		o += c.length;
	}
	return out;
}

async function* parts(list: StreamPart[]): AsyncGenerator<StreamPart> {
	for (const p of list) yield p;
}

describe("streamZip", () => {
	it("round-trips buffered, sync-iterable, and async-iterable parts through openZip", async () => {
		const rows = Array.from({ length: 2000 }, (_, i) => `<row r="${i + 1}">data ${i}</row>`);
		async function* asyncChunks(): AsyncGenerator<Uint8Array> {
			for (const r of rows) yield enc.encode(r);
		}
		const bytes = await collect(
			streamZip(
				parts([
					{ name: "a.txt", data: enc.encode("hello world") },
					{ name: "sync.xml", data: [enc.encode("one"), enc.encode("two")] }, // sync iterable
					{ name: "xl/big.xml", data: asyncChunks() }, // async stream (the worksheet shape)
					{ name: "empty.txt", data: enc.encode("") },
				]),
			),
		);
		const zip = openZip(bytes);
		expect(dec.decode(await zip.read("a.txt"))).toBe("hello world");
		expect(dec.decode(await zip.read("sync.xml"))).toBe("onetwo");
		expect(dec.decode(await zip.read("xl/big.xml"))).toBe(rows.join(""));
		expect(dec.decode(await zip.read("empty.txt"))).toBe("");
	});

	it("rejects a duplicate name and a directory-placeholder name", async () => {
		await expect(
			collect(
				streamZip(
					parts([
						{ name: "x", data: enc.encode("a") },
						{ name: "x", data: enc.encode("b") },
					]),
				),
			),
		).rejects.toThrow(/duplicate/);
		await expect(
			collect(streamZip(parts([{ name: "dir/", data: enc.encode("a") }]))),
		).rejects.toThrow(/directory placeholder/);
	});

	it("propagates an error from an async data source instead of hanging", async () => {
		async function* boom(): AsyncGenerator<Uint8Array> {
			yield enc.encode("<partial>");
			throw new Error("row source failed");
		}
		await expect(collect(streamZip(parts([{ name: "x.xml", data: boom() }])))).rejects.toThrow(
			/row source failed/,
		);
	});
});
