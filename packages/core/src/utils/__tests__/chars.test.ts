import { describe, expect, it } from "vitest"
import { isWhitespace } from "../chars"

describe("isWhitespace", () => {
	it("is true for the four XML whitespace characters", () => {
		for (const ch of [" ", "\t", "\n", "\r"]) {
			expect(isWhitespace(ch)).toBe(true)
		}
	})

	it("is false for non-XML-whitespace, including form feed and vertical tab", () => {
		for (const ch of ["a", "0", "<", "\f", "\v", " ", ""]) {
			expect(isWhitespace(ch)).toBe(false)
		}
	})

	it("is false for undefined (end of input)", () => {
		expect(isWhitespace(undefined)).toBe(false)
	})
})
