import { XlsxError } from "../errors";

// Raw DEFLATE decompression using the platform Compression Streams API — available on
// Node >= 18, Deno, Bun, modern browsers, and Cloudflare Workers. This is why openjsxl
// needs no zip/inflate dependency: the runtime already ships one.
//
// `maxBytes` bounds the output: decompression is aborted and a typed `part-too-large` error
// thrown once the inflated size would exceed it — the abort happens DURING streaming, so a
// decompression bomb never fully materializes. The caller (openZip) passes the tightest of the
// declared size, the absolute per-part ceiling, and the compression-ratio limit (F9.7).

// A distinct XlsxError so openZip can tell a legitimate bomb/size abort (surface as-is) from a
// genuinely corrupt deflate stream (wrap as corrupt-zip).
function tooLarge(maxBytes: number): XlsxError {
	return new XlsxError(
		"part-too-large",
		`inflated output exceeds the ${maxBytes}-byte limit (zip-bomb / size guard)`,
	);
}

export async function inflateRaw(
	data: Uint8Array,
	maxBytes = Number.POSITIVE_INFINITY,
): Promise<Uint8Array> {
	// The Blob accepts our bytes at runtime; the cast sidesteps a TS lib / @types/node
	// disagreement over whether the backing buffer is an ArrayBuffer or SharedArrayBuffer.
	const blob = new Blob([data as BlobPart]);
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader();

	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.byteLength;
		if (total > maxBytes) {
			await reader.cancel();
			throw tooLarge(maxBytes);
		}
		chunks.push(value);
	}

	const out = new Uint8Array(total);
	let offset = 0;
	for (const chunk of chunks) {
		out.set(chunk, offset);
		offset += chunk.byteLength;
	}
	return out;
}

// Streaming variant: yield decompressed chunks instead of one buffer, so a large part can be
// consumed without ever materializing it whole. `maxBytes` bounds the total output (a bomb
// cap); the stream is cancelled if the consumer stops early.
export async function* inflateRawStream(
	data: Uint8Array,
	maxBytes = Number.POSITIVE_INFINITY,
): AsyncGenerator<Uint8Array> {
	const blob = new Blob([data as BlobPart]);
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader();

	let total = 0;
	try {
		for (;;) {
			const { done, value } = await reader.read();
			if (done) break;
			total += value.byteLength;
			if (total > maxBytes) {
				throw tooLarge(maxBytes);
			}
			yield value;
		}
	} finally {
		await reader.cancel().catch(() => {});
	}
}
