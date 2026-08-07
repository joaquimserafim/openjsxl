import { XlsxError } from "../errors";

// Raw DEFLATE decompression using the platform Compression Streams API — available on
// Node >= 18, Deno, Bun, modern browsers, and Cloudflare Workers. This is why openjsxl
// needs no zip/inflate dependency: the runtime already ships one.
//
// `maxBytes` bounds the output: decompression is aborted and a typed `part-too-large` error
// thrown once the inflated size would exceed it — the abort happens DURING streaming, so a
// decompression bomb never fully materializes. The caller (openZip) passes the tightest of the
// declared size, the absolute per-part ceiling, and the compression-ratio limit (F9.7).

// A distinct XlsxError keeps a legitimate bomb/size abort separate from a corrupt deflate stream.
function tooLarge(maxBytes: number): XlsxError {
	return new XlsxError(
		"part-too-large",
		`inflated output exceeds the ${maxBytes}-byte limit (zip-bomb / size guard)`,
	);
}

function inflateFailure(entryName: string | undefined, cause: unknown): XlsxError {
	if (cause instanceof XlsxError) return cause;
	const target = entryName === undefined ? "entry" : entryName;
	return new XlsxError("corrupt-zip", `corrupt zip: failed to inflate ${target}`, { cause });
}

async function readInflatedChunk(
	reader: ReadableStreamDefaultReader<Uint8Array>,
	entryName: string | undefined,
): Promise<ReadableStreamReadResult<Uint8Array>> {
	return reader.read().catch((cause: unknown) => {
		throw inflateFailure(entryName, cause);
	});
}

export async function inflateRaw(
	data: Uint8Array,
	maxBytes = Number.POSITIVE_INFINITY,
	entryName?: string,
): Promise<Uint8Array> {
	// The Blob accepts our bytes at runtime; the cast sidesteps a TS lib / @types/node
	// disagreement over whether the backing buffer is an ArrayBuffer or SharedArrayBuffer.
	const blob = new Blob([data as BlobPart]);
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader();

	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await readInflatedChunk(reader, entryName);
		if (done) break;
		total += value.byteLength;
		if (total > maxBytes) {
			// Cleanup failure must not replace the deterministic size-limit error already observed.
			await reader.cancel().catch(() => {});
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
	entryName?: string,
): AsyncGenerator<Uint8Array> {
	const blob = new Blob([data as BlobPart]);
	const reader = blob.stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader();

	let total = 0;
	try {
		for (;;) {
			const { done, value } = await readInflatedChunk(reader, entryName);
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
