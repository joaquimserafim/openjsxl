import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { MAX_FORMULA_LEN } from "../../ooxml/formula";
import { MAX_NAME_LEN } from "../../ooxml/name";
import type { DefinedName } from "../../ooxml/workbook";
import { openXlsx } from "../../reader/workbook";
import { openZip } from "../../zip";
import { dedupeDefinedNames, workbookToInput } from "../from-workbook";
import { streamXlsx } from "../stream";
import type { WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

// F10.1 — workbook defined names: write + bridge carry. Everything written must re-read through
// Workbook.definedNames verbatim (shared model) and cross the bridge; a names-free workbook must keep
// its exact pre-F10.1 workbook.xml bytes.

const decoder = new TextDecoder();

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		total += value.length;
	}
	const out = new Uint8Array(total);
	let o = 0;
	for (const c of chunks) {
		out.set(c, o);
		o += c.length;
	}
	return out;
}

const workbookXmlOf = async (bytes: Uint8Array): Promise<string> =>
	decoder.decode(await openZip(bytes).read("xl/workbook.xml"));

// Capture writeXlsx's typed rejection so a test can assert its code + message.
async function writeErr(wb: unknown): Promise<XlsxError> {
	try {
		await writeXlsx(wb as WorkbookInput);
	} catch (e) {
		if (e instanceof XlsxError) return e;
		throw e;
	}
	throw new Error("expected writeXlsx to reject");
}

describe("writeXlsx — <definedNames> emission", () => {
	it("emits names in the CT_Workbook slot, right after </sheets>", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Sheet1", rows: [["a"]] }],
			definedNames: [{ name: "Total", refersTo: "Sheet1!$A$1:$A$9" }],
		});
		const xml = await workbookXmlOf(bytes);
		expect(xml).toContain(
			'</sheets><definedNames><definedName name="Total">Sheet1!$A$1:$A$9</definedName></definedNames></workbook>',
		);
	});

	it("emits localSheetId and hidden attributes (order: name, localSheetId, hidden)", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Sheet1", rows: [["a"]] }],
			definedNames: [
				{ name: "Local", refersTo: "Sheet1!$B$2", localSheetId: 0, hidden: true },
			],
		});
		expect(await workbookXmlOf(bytes)).toContain(
			'<definedName name="Local" localSheetId="0" hidden="1">Sheet1!$B$2</definedName>',
		);
	});

	it("escapes the name attribute and the refersTo element text", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Sheet1", rows: [["a"]] }],
			// A string-constant name whose refersTo contains markup characters.
			definedNames: [{ name: "Greeting", refersTo: '"a<b&c"' }],
		});
		const xml = await workbookXmlOf(bytes);
		expect(xml).toContain('<definedName name="Greeting">"a&lt;b&amp;c"</definedName>');
	});

	it("emits nothing for hidden:false (default) — no hidden attribute", async () => {
		const bytes = await writeXlsx({
			sheets: [{ name: "Sheet1", rows: [["a"]] }],
			definedNames: [{ name: "Vis", refersTo: "Sheet1!$A$1", hidden: false }],
		});
		expect(await workbookXmlOf(bytes)).toContain(
			'<definedName name="Vis">Sheet1!$A$1</definedName>',
		);
	});
});

describe("writeXlsx — byte-identity when no names are used", () => {
	const base: WorkbookInput = { sheets: [{ name: "Sheet1", rows: [["a", 1]] }] };

	it("absent, undefined, and empty definedNames all produce identical workbook.xml with no <definedNames>", async () => {
		const none = await workbookXmlOf(await writeXlsx(base));
		// An untyped JS caller can pass definedNames: undefined explicitly (exactOptionalPropertyTypes
		// forbids it at the type level); the writer treats it as absent.
		const undefInput: unknown = { ...base, definedNames: undefined };
		const undef = await workbookXmlOf(await writeXlsx(undefInput as WorkbookInput));
		const empty = await workbookXmlOf(await writeXlsx({ ...base, definedNames: [] }));
		expect(none).not.toContain("<definedNames");
		expect(undef).toBe(none);
		expect(empty).toBe(none);
	});
});

describe("writeXlsx — round-trip through the reader (shared model)", () => {
	it("names re-read verbatim through Workbook.definedNames, in order", async () => {
		const definedNames: DefinedName[] = [
			{ name: "Total", refersTo: "Sheet1!$A$1:$A$9" },
			{ name: "Local", refersTo: "Sheet1!$B$2", localSheetId: 0, hidden: true },
			{ name: "_xlnm.Print_Area", refersTo: "Sheet1!$A$1:$C$3", localSheetId: 0 },
		];
		const wb = await openXlsx(
			await writeXlsx({ sheets: [{ name: "Sheet1", rows: [["a"]] }], definedNames }),
		);
		expect(wb.definedNames).toEqual(definedNames);
	});

	it("round-trips a name colliding with the _xHHHH_ escape shape (ST_Xstring @name, F9.6 parity)", async () => {
		const definedNames = [{ name: "_x0041_", refersTo: "Sheet1!$A$1" }];
		const bytes = await writeXlsx({
			sheets: [{ name: "Sheet1", rows: [["a"]] }],
			definedNames,
		});
		// Emitted ENCODED so Excel/openpyxl decode back to the literal name (not the character "A").
		expect(await workbookXmlOf(bytes)).toContain(
			'<definedName name="_x005F_x0041_">Sheet1!$A$1</definedName>',
		);
		// And our own reader decodes it back to the true name.
		expect((await openXlsx(bytes)).definedNames).toEqual(definedNames);
	});

	it("streamXlsx carries the same names (reader-equivalent to writeXlsx)", async () => {
		const definedNames: DefinedName[] = [
			{ name: "Total", refersTo: "Sheet1!$A$1" },
			{ name: "Scoped", refersTo: "Sheet1!$B$1", localSheetId: 0 },
		];
		const wb = await openXlsx(
			await drain(streamXlsx({ sheets: [{ name: "Sheet1", rows: [["a"]] }], definedNames })),
		);
		expect(wb.definedNames).toEqual(definedNames);
	});
});

describe("writeXlsx — validation rejects (typed invalid-input)", () => {
	const sheets = [{ name: "Sheet1", rows: [["a"]] }];
	const withNames = (definedNames: unknown): unknown => ({ sheets, definedNames });

	it("rejects a non-array definedNames", async () => {
		const err = await writeErr(withNames({ name: "X", refersTo: "1" }));
		expect(err.code).toBe("invalid-input");
		expect(err.message).toContain("definedNames must be an array");
	});

	it("rejects a non-object entry and an unknown property", async () => {
		expect((await writeErr(withNames(["nope"]))).message).toContain("must be an object");
		expect(
			(await writeErr(withNames([{ name: "X", refersTo: "1", color: "red" }]))).message,
		).toContain('unknown property "color"');
	});

	it("rejects an illegal name, naming the broken rule", async () => {
		expect((await writeErr(withNames([{ name: "My Name", refersTo: "1" }]))).message).toContain(
			"whitespace",
		);
		expect((await writeErr(withNames([{ name: "A1", refersTo: "1" }]))).message).toContain(
			"cell reference",
		);
		expect(
			(await writeErr(withNames([{ name: "_xlnm.Nope", refersTo: "1" }]))).message,
		).toContain("_xlnm.");
		expect((await writeErr(withNames([{ name: 42, refersTo: "1" }]))).message).toContain(
			"name must be a string",
		);
	});

	it("rejects a bad refersTo (non-string, empty, oversized, leading =, XML-unsafe)", async () => {
		expect((await writeErr(withNames([{ name: "X", refersTo: 1 }]))).message).toContain(
			"refersTo must be a string",
		);
		expect((await writeErr(withNames([{ name: "X", refersTo: "" }]))).message).toContain(
			"must not be empty",
		);
		expect(
			(await writeErr(withNames([{ name: "X", refersTo: "a".repeat(MAX_FORMULA_LEN + 1) }])))
				.message,
		).toContain("character limit");
		expect((await writeErr(withNames([{ name: "X", refersTo: "=A1" }]))).message).toContain(
			"stored form",
		);
		expect((await writeErr(withNames([{ name: "X", refersTo: "A\x01" }]))).message).toContain(
			"not allowed in XML",
		);
	});

	it("rejects a localSheetId that is not an integer index of an existing sheet", async () => {
		expect(
			(await writeErr(withNames([{ name: "X", refersTo: "1", localSheetId: 5 }]))).message,
		).toContain("existing sheet");
		expect(
			(await writeErr(withNames([{ name: "X", refersTo: "1", localSheetId: 1.5 }]))).message,
		).toContain("existing sheet");
	});

	it("rejects a non-boolean hidden", async () => {
		expect(
			(await writeErr(withNames([{ name: "X", refersTo: "1", hidden: "yes" }]))).message,
		).toContain("hidden must be a boolean");
	});

	it("rejects a duplicate name in the same scope (case-insensitive) but allows different scopes", async () => {
		// Same global scope, case-insensitive collision → reject.
		const dup = await writeErr(
			withNames([
				{ name: "Foo", refersTo: "1" },
				{ name: "foo", refersTo: "2" },
			]),
		);
		expect(dup.message).toContain("duplicate name");
		// Global + sheet-local of the same spelling is legal (different scope).
		const ok = await writeXlsx(
			withNames([
				{ name: "Foo", refersTo: "Sheet1!$A$1" },
				{ name: "Foo", refersTo: "Sheet1!$B$1", localSheetId: 0 },
			]) as WorkbookInput,
		);
		expect((await openXlsx(ok)).definedNames).toHaveLength(2);
	});

	it("accepts a name exactly at the length limit", async () => {
		const name = `_${"a".repeat(MAX_NAME_LEN - 1)}`;
		const wb = await openXlsx(
			await writeXlsx({ sheets, definedNames: [{ name, refersTo: "Sheet1!$A$1" }] }),
		);
		expect(wb.definedNames[0]?.name).toBe(name);
	});
});

describe("bridge — workbookToInput carries defined names", () => {
	it("round-trips names read from one written file into another", async () => {
		const definedNames: DefinedName[] = [
			{ name: "Total", refersTo: "Sheet1!$A$1:$A$9" },
			{ name: "Local", refersTo: "Sheet1!$B$2", localSheetId: 0, hidden: true },
		];
		const original = await openXlsx(
			await writeXlsx({ sheets: [{ name: "Sheet1", rows: [["a"]] }], definedNames }),
		);
		const rewritten = await openXlsx(await writeXlsx(await workbookToInput(original)));
		expect(rewritten.definedNames).toEqual(definedNames);
	});

	it("omits definedNames from the bridge output when the source has none (byte-identity path)", async () => {
		const source = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] }));
		const input = await workbookToInput(source);
		expect(input.definedNames).toBeUndefined();
	});
});

describe("dedupeDefinedNames", () => {
	it("returns the same array reference when there are no duplicates", () => {
		const names: DefinedName[] = [
			{ name: "A", refersTo: "1" },
			{ name: "B", refersTo: "2", localSheetId: 0 },
		];
		expect(dedupeDefinedNames(names)).toBe(names);
	});

	it("drops a later per-scope case-insensitive duplicate, keeping the first", () => {
		const out = dedupeDefinedNames([
			{ name: "Foo", refersTo: "first" },
			{ name: "foo", refersTo: "second" }, // same global scope, case-insensitive collision
			{ name: "Foo", refersTo: "scoped", localSheetId: 0 }, // different scope — kept
		]);
		expect(out).toEqual([
			{ name: "Foo", refersTo: "first" },
			{ name: "Foo", refersTo: "scoped", localSheetId: 0 },
		]);
	});
});
