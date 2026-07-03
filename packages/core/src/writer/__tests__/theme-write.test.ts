import { loadFixture } from "@openjsxl/fixtures"
import { describe, expect, it } from "vitest"
import { XlsxError } from "../../errors"
import { openXlsx } from "../../reader/workbook"
import { openZip } from "../../zip"
import { workbookToInput } from "../from-workbook"
import { DEFAULT_THEME_XML } from "../theme"
import { writeXlsx } from "../workbook"

// F5.3 — theme carry. A written style that uses a theme color needs a theme part; the writer emits
// the workbook's carried theme when present (so custom-theme files keep their colors on rewrite) or
// the built-in Office theme otherwise (keeping pre-F5.3 bytes for everything else).

const decoder = new TextDecoder()
const themePart = async (bytes: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(bytes).read("xl/theme/theme1.xml"))

// A minimal workbook whose one cell uses a theme color — enough to force a theme part.
const themedInput = (themeXml?: string) => ({
	sheets: [{ name: "S", rows: [[{ value: "x", style: { font: { color: { theme: 4 } } } }]] }],
	...(themeXml !== undefined ? { themeXml } : {}),
})

describe("writeXlsx — theme carry (custom theme survives the round trip)", () => {
	it("re-emits the source theme byte-identically and resolves the same color", async () => {
		const src = await loadFixture("openpyxl-customtheme.xlsx") // accent1 recolored FF0000
		const before = await openXlsx(src)
		const rewritten = await writeXlsx(await workbookToInput(before))
		// The theme part is carried verbatim...
		expect(await themePart(rewritten)).toBe(await themePart(src))
		// ...so the themed cell still resolves against the CUSTOM accent1, not the default.
		const after = await openXlsx(rewritten)
		const color = after.sheet("Themed").style("B2")?.font?.color as { theme: number }
		expect(after.resolveColor(color)).toBe("FFFF0000")
	})
})

describe("writeXlsx — built-in theme fallback", () => {
	it("emits the default Office theme when no theme is carried", async () => {
		const bytes = await writeXlsx(themedInput())
		expect(await themePart(bytes)).toBe(DEFAULT_THEME_XML)
		// resolveColor on our own output uses the default accent1 (4F81BD).
		const wb = await openXlsx(bytes)
		expect(wb.resolveColor({ theme: 4 })).toBe("FF4F81BD")
	})

	it("ignores a carried theme when no written style needs a theme part", async () => {
		// No theme-colored cell → no theme part → themeXml is irrelevant (and pre-F5.3 bytes hold).
		const bytes = await writeXlsx({
			sheets: [{ name: "S", rows: [["plain"]] }],
			themeXml: "<x/>",
		})
		expect(openZip(bytes).has("xl/theme/theme1.xml")).toBe(false)
	})

	it("degrades a present-but-empty theme part to the built-in theme (review regression)", async () => {
		// A truncated/corrupt producer can leave xl/theme/theme1.xml present but 0 bytes. The reader
		// treats that as no theme, so the bridge doesn't carry "" into the writer's non-empty check —
		// which used to throw invalid-input on a file the reader had accepted.
		const { writeZip } = await import("../zip")
		const themed = await writeXlsx(themedInput()) // contains a real xl/theme/theme1.xml
		const zip = openZip(themed)
		const parts: { name: string; data: Uint8Array }[] = []
		for (const name of zip.entries.keys()) {
			parts.push({
				name,
				data: name === "xl/theme/theme1.xml" ? new Uint8Array(0) : await zip.read(name),
			})
		}
		const withEmptyTheme = await writeZip(parts)
		const wb = await openXlsx(withEmptyTheme)
		expect(wb.themeXml).toBeUndefined() // empty part → treated as absent
		// The rewrite succeeds (no throw) and falls back to the built-in theme.
		const rewritten = await writeXlsx(await workbookToInput(wb))
		expect(await themePart(rewritten)).toBe(DEFAULT_THEME_XML)
	})
})

describe("writeXlsx — themeXml validation", () => {
	const reject = async (themeXml: unknown, pattern: RegExp): Promise<void> => {
		const err = await writeXlsx({
			// biome-ignore lint/suspicious/noExplicitAny: exercising input the types forbid
			...themedInput(themeXml as any),
		}).then(
			() => undefined,
			(e) => e,
		)
		expect(err).toBeInstanceOf(XlsxError)
		expect((err as XlsxError).code).toBe("invalid-input")
		expect((err as XlsxError).message).toMatch(pattern)
	}

	it("rejects an empty or XML-unsafe carried theme", async () => {
		await reject("", /themeXml must be a non-empty string/)
		await reject(
			`the${String.fromCharCode(1)}me`,
			/themeXml contains a character not allowed in XML/,
		)
	})
})
