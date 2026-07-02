import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { XlsxError } from "../errors"
import { openXlsx, streamSheetRows } from "../reader/workbook"
import { openZip } from "../zip"

// Every file-level failure surfaces as an XlsxError with a discriminating `code` (F2.4a).
// Catch the thrown/rejected value so we can assert its type and code, not just its message.
async function caught(fn: () => unknown): Promise<unknown> {
	try {
		await fn()
	} catch (e) {
		return e
	}
	throw new Error("expected the call to throw")
}

describe("XlsxError", () => {
	it("is an Error subclass with a stable name and a code", async () => {
		const e = await caught(() => openZip(new Uint8Array([1, 2, 3, 4])))
		expect(e).toBeInstanceOf(Error)
		expect(e).toBeInstanceOf(XlsxError)
		expect((e as XlsxError).name).toBe("XlsxError")
		expect((e as XlsxError).code).toBe("not-a-zip")
	})

	it("codes a missing zip entry as missing-part", async () => {
		const zip = openZip(await loadFixture("basic.xlsx"))
		const e = await caught(() => zip.read("nope.xml"))
		expect((e as XlsxError).code).toBe("missing-part")
	})

	it("codes an unknown sheet name as no-such-sheet (random access)", async () => {
		const wb = await openXlsx(await loadFixture("basic.xlsx"))
		const e = await caught(() => wb.sheet("Nope"))
		expect(e).toBeInstanceOf(XlsxError)
		expect((e as XlsxError).code).toBe("no-such-sheet")
	})

	it("codes an unknown sheet name as no-such-sheet (streaming)", async () => {
		const bytes = await loadFixture("basic.xlsx")
		const e = await caught(async () => {
			for await (const _row of streamSheetRows(bytes, "Nope")) break
		})
		expect((e as XlsxError).code).toBe("no-such-sheet")
	})

	it("codes a package with no officeDocument relationship as not-xlsx", async () => {
		const bytes = await loadFixture("broken-no-officedoc.xlsx")
		const e = await caught(() => openXlsx(bytes))
		expect((e as XlsxError).code).toBe("not-xlsx")
	})

	it("codes a package missing a required part as missing-part", async () => {
		const bytes = await loadFixture("broken-no-workbook.xlsx")
		const e = await caught(() => openXlsx(bytes))
		expect((e as XlsxError).code).toBe("missing-part")
	})
})
