// Raw DEFLATE decompression using the platform Compression Streams API — available on
// Node >= 18, Deno, Bun, modern browsers, and Cloudflare Workers. This is why openjsxl
// needs no zip/inflate dependency: the runtime already ships one.
//
// `maxBytes` bounds the output: decompression is aborted and an error thrown once the
// inflated size would exceed it. Callers pass the entry's declared uncompressed size so a
// malformed or hostile stream can't expand without bound (a decompression bomb).

export async function inflateRaw(
	data: Uint8Array,
	maxBytes = Number.POSITIVE_INFINITY,
): Promise<Uint8Array> {
	// The Blob accepts our bytes at runtime; the cast sidesteps a TS lib / @types/node
	// disagreement over whether the backing buffer is an ArrayBuffer or SharedArrayBuffer.
	const blob = new Blob([data as BlobPart])
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader()

	const chunks: Uint8Array[] = []
	let total = 0
	for (;;) {
		const { done, value } = await reader.read()
		if (done) break
		total += value.byteLength
		if (total > maxBytes) {
			await reader.cancel()
			throw new Error(`inflated output exceeds the expected ${maxBytes} bytes`)
		}
		chunks.push(value)
	}

	const out = new Uint8Array(total)
	let offset = 0
	for (const chunk of chunks) {
		out.set(chunk, offset)
		offset += chunk.byteLength
	}
	return out
}

// Streaming variant: yield decompressed chunks instead of one buffer, so a large part can be
// consumed without ever materializing it whole. `maxBytes` bounds the total output (a bomb
// cap); the stream is cancelled if the consumer stops early.
export async function* inflateRawStream(
	data: Uint8Array,
	maxBytes = Number.POSITIVE_INFINITY,
): AsyncGenerator<Uint8Array> {
	const blob = new Blob([data as BlobPart])
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader()

	let total = 0
	try {
		for (;;) {
			const { done, value } = await reader.read()
			if (done) break
			total += value.byteLength
			if (total > maxBytes) {
				throw new Error(`inflated output exceeds the expected ${maxBytes} bytes`)
			}
			yield value
		}
	} finally {
		await reader.cancel().catch(() => {})
	}
}
