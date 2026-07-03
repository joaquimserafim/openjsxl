import { describe, expect, it } from "vitest";
import { openZip } from "../../zip";
import { MAX_ENTRIES as STREAM_MAX_ENTRIES, type StreamPart, streamZip } from "../stream-zip";
import { MAX_ENTRIES as BUFFERED_MAX_ENTRIES } from "../zip";

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
	it("caps entries below the ZIP64 sentinel, same bound in both writers (M5-analysis regression)", () => {
		// A 65 536-part archive is too slow to build in a unit test, so pin the shared constant: the
		// u16 EOCD entry count 0xffff is the ZIP64 sentinel — writing exactly 65 535 entries would
		// send conforming readers hunting for a ZIP64 record that doesn't exist. Max writable: 0xfffe.
		expect(STREAM_MAX_ENTRIES).toBe(0xfffe);
		expect(BUFFERED_MAX_ENTRIES).toBe(0xfffe);
	});

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
