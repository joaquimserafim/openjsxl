// Raw DEFLATE compression via the platform Compression Streams API — the mirror image of
// zip/inflate.ts. Leaning on the runtime's CompressionStream (Node >= 18, Deno, Bun, modern
// browsers, Cloudflare Workers) is what keeps the writer dependency-free, exactly as
// DecompressionStream does for the reader.
//
// `'deflate-raw'` emits a bare DEFLATE stream (RFC 1951) with no zlib header or Adler-32 trailer —
// which is precisely what a ZIP entry with compression method 8 stores. (`'deflate'`, by contrast,
// would add the zlib wrapper the reader's inflateRaw does not expect.)

export async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	// The Blob accepts our bytes at runtime; the cast sidesteps a TS lib / @types/node
	// disagreement over whether the backing buffer is an ArrayBuffer or SharedArrayBuffer.
	const blob = new Blob([data as BlobPart]);
	const reader = blob.stream().pipeThrough(new CompressionStream("deflate-raw")).getReader();

	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.byteLength;
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
