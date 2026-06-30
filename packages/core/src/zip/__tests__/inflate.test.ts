import { describe, expect, it } from 'vitest'
import { inflateRaw } from '../inflate'

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	const stream = new Blob([data as unknown as BlobPart])
		.stream()
		.pipeThrough(new CompressionStream('deflate-raw'))
	return new Uint8Array(await new Response(stream).arrayBuffer())
}

describe('inflateRaw', () => {
	it('round-trips text compressed with the platform deflate-raw', async () => {
		const text = 'hello openjsxl — '.repeat(2000)
		const original = new TextEncoder().encode(text)
		const restored = await inflateRaw(await deflateRaw(original))
		expect(new TextDecoder().decode(restored)).toBe(text)
	})

	it('round-trips binary data across a range of sizes', async () => {
		for (const size of [0, 1, 100, 65536]) {
			const original = new Uint8Array(size).map((_, i) => i % 256)
			const restored = await inflateRaw(await deflateRaw(original))
			expect(restored).toEqual(original)
		}
	})
})
