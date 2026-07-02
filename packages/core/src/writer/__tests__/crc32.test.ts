import { describe, expect, it } from "vitest"
import { crc32 } from "../crc32"

// CRC-32 (IEEE) has well-known check values; if our table or accumulation is off, these fail.
// The `"123456789"` vector (0xCBF43926) is the canonical CRC-32 check constant.

const enc = new TextEncoder()

describe("crc32", () => {
	it("is 0 for empty input", () => {
		expect(crc32(new Uint8Array(0))).toBe(0)
	})

	it("matches the canonical check vector", () => {
		expect(crc32(enc.encode("123456789"))).toBe(0xcbf43926)
	})

	it("matches known text vectors", () => {
		expect(crc32(enc.encode("The quick brown fox jumps over the lazy dog"))).toBe(0x414fa339)
		expect(crc32(enc.encode("hello"))).toBe(0x3610a686)
	})

	it("returns an unsigned 32-bit integer", () => {
		// Raw bytes whose CRC has the high bit set — the final `>>> 0` must keep the result
		// positive rather than letting JS's signed `^` produce a negative number.
		const crc = crc32(Uint8Array.from([0xff, 0xff, 0xff, 0xff]))
		expect(crc).toBeGreaterThanOrEqual(0)
		expect(crc).toBeLessThanOrEqual(0xffffffff)
		expect(Number.isInteger(crc)).toBe(true)
	})
})
