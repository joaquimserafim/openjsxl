import { describe, expect, it } from "vitest"
import { localName } from "../xml-names"

describe("localName", () => {
	it("strips a namespace prefix", () => {
		expect(localName("r:id")).toBe("id")
		expect(localName("xml:space")).toBe("space")
	})

	it("returns unprefixed names unchanged", () => {
		expect(localName("si")).toBe("si")
		expect(localName("")).toBe("")
	})

	it("splits on the first colon only", () => {
		expect(localName("a:b:c")).toBe("b:c")
	})
})
