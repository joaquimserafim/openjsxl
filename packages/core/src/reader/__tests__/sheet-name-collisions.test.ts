import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { workbookToInput } from "../../writer/from-workbook";
import { writeXlsx } from "../../writer/workbook";
import { writeZip } from "../../writer/zip";
import { openXlsx, streamSheetRows } from "../workbook";

// F11.40(a), owner call D17: two tabs whose decoded names match ignoring case share one name. The
// reader keeps the first such tab that loads and leaves each later one out, with its sheet-scoped
// defined names. The comparison is `toLowerCase`, the writer's own duplicate rule.

const enc = new TextEncoder();
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";

const KELVIN = "\u212A"; // KELVIN SIGN; lower-cases to "k"
const DOTTED_I = "\u0130"; // LATIN CAPITAL LETTER I WITH DOT ABOVE; lower-cases to "i" + U+0307
const I_DOT = "i\u0307"; // "i" + COMBINING DOT ABOVE
const DOTLESS_I = "\u0131"; // LATIN SMALL LETTER DOTLESS I

/**
 * A workbook whose tab `i` has the wire name `names[i]`, written as openpyxl 3.1.5 writes it (an
 * escape-shaped title such as `_x0041_` goes out raw), holds `i + 1` in A1, and owns the sheet-scoped
 * name `Scope<i>` = `i`. A global `Total` = 42 comes first.
 */
function tabsWorkbook(names: readonly string[]): Promise<Uint8Array> {
	const worksheet = (value: number) =>
		`<?xml version="1.0"?><worksheet xmlns="${NS_MAIN}"><sheetData>` +
		`<row r="1"><c r="A1"><v>${value}</v></c></row></sheetData></worksheet>`;
	const workbook =
		`<?xml version="1.0"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets>` +
		names
			.map(
				(name, i) =>
					`<sheet name="${name}" sheetId="${i + 1}" state="visible" r:id="rId${i + 1}" />`,
			)
			.join("") +
		`</sheets><definedNames><definedName name="Total">42</definedName>` +
		names
			.map(
				(_name, i) =>
					`<definedName name="Scope${i}" localSheetId="${i}">${i}</definedName>`,
			)
			.join("") +
		`</definedNames></workbook>`;
	const wbRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		names
			.map(
				(_name, i) =>
					`<Relationship Id="rId${i + 1}" Type="${NS_R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
			)
			.join("") +
		`</Relationships>`;
	const rootRels =
		`<?xml version="1.0"?><Relationships xmlns="${NS_PKG}">` +
		`<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	const types =
		`<?xml version="1.0"?><Types xmlns="${CT}">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
		`</Types>`;
	return writeZip([
		{ name: "[Content_Types].xml", data: enc.encode(types) },
		{ name: "_rels/.rels", data: enc.encode(rootRels) },
		{ name: "xl/workbook.xml", data: enc.encode(workbook) },
		{ name: "xl/_rels/workbook.xml.rels", data: enc.encode(wbRels) },
		...names.map((_name, i) => ({
			name: `xl/worksheets/sheet${i + 1}.xml`,
			data: enc.encode(worksheet(i + 1)),
		})),
	]);
}

/** Each loaded tab as `name=A1`, in tab order. */
async function loadedTabs(bytes: Uint8Array): Promise<string[]> {
	const wb = await openXlsx(bytes);
	return wb.sheets.map((s) => `${s.name}=${String(wb.sheet(s.name).cell("A1")?.value)}`);
}

/** The first streamed A1 value of the named tab (or of the first tab). */
async function streamedA1(bytes: Uint8Array, name?: string): Promise<unknown> {
	for await (const row of streamSheetRows(bytes, name)) return row.cells[0]?.value;
	return undefined;
}

const TOTAL = { name: "Total", refersTo: "42" };
const scope = (raw: number, loaded = raw) => ({
	name: `Scope${raw}`,
	refersTo: String(raw),
	localSheetId: loaded,
});

describe("reader — tabs whose names match ignoring case keep the first (F11.40(a), D17)", () => {
	it("keeps A of a raw case-only pair [A, a] and leaves the later a out, with its scoped name", async () => {
		const bytes = await tabsWorkbook(["A", "a"]);
		const wb = await openXlsx(bytes);
		expect(await loadedTabs(bytes)).toEqual(["A=1"]);
		expect(() => wb.sheet("a")).toThrow(/no sheet named "a"; available: A$/);
		expect(wb.definedNames).toEqual([TOTAL, scope(0)]);
		expect(await streamedA1(bytes)).toBe(1);
		expect(await streamedA1(bytes, "A")).toBe(1);
		await expect(streamedA1(bytes, "a")).rejects.toMatchObject({ code: "no-such-sheet" });
	});

	it("keeps whichever spelling comes first: [a, A] keeps a", async () => {
		const bytes = await tabsWorkbook(["a", "A"]);
		expect(await loadedTabs(bytes)).toEqual(["a=1"]);
		expect((await openXlsx(bytes)).definedNames).toEqual([TOTAL, scope(0)]);
		await expect(streamedA1(bytes, "A")).rejects.toMatchObject({ code: "no-such-sheet" });
	});

	it("reads a raw _x0041_ tab, as openpyxl 3.1.5 writes one, as A and leaves the real A tab out", async () => {
		// openpyxl 3.1.5 writes the title `_x0041_` unescaped (this package copies its <sheet> shape);
		// OOXML and Excel decode it as `A`.
		const bytes = await tabsWorkbook(["_x0041_", "A"]);
		const wb = await openXlsx(bytes);
		expect(await loadedTabs(bytes)).toEqual(["A=1"]);
		expect(wb.definedNames).toEqual([TOTAL, scope(0)]);
		expect(await streamedA1(bytes, "A")).toBe(1);
		await expect(streamedA1(bytes, "_x0041_")).rejects.toMatchObject({ code: "no-such-sheet" });
	});

	it.each([
		["Kelvin sign then K", [KELVIN, "K"], KELVIN],
		["K then Kelvin sign", ["K", KELVIN], "K"],
		["Kelvin sign then k", [KELVIN, "k"], KELVIN],
		["raw _x212A_ then K", ["_x212A_", "K"], KELVIN],
		["dotted I then i + combining dot", [DOTTED_I, I_DOT], DOTTED_I],
		["i + combining dot then dotted I", [I_DOT, DOTTED_I], I_DOT],
	] as const)("treats %s as one name and keeps the first", async (_label, names, kept) => {
		const bytes = await tabsWorkbook(names);
		expect(await loadedTabs(bytes)).toEqual([`${kept}=1`]);
		expect((await openXlsx(bytes)).definedNames).toEqual([TOTAL, scope(0)]);
		expect(await streamedA1(bytes, kept)).toBe(1);
	});

	it.each([
		["dotted I and plain i", [DOTTED_I, "i"]],
		["I and dotless i", ["I", DOTLESS_I]],
	] as const)("keeps both tabs of %s, which lower-case apart", async (_label, names) => {
		const bytes = await tabsWorkbook(names);
		expect(await loadedTabs(bytes)).toEqual([`${names[0]}=1`, `${names[1]}=2`]);
		expect((await openXlsx(bytes)).definedNames).toEqual([TOTAL, scope(0), scope(1)]);
		expect(await streamedA1(bytes, names[1])).toBe(2);
	});

	it("remaps the scoped names of the tabs after a dropped one: [X, A, a, Y]", async () => {
		const bytes = await tabsWorkbook(["X", "A", "a", "Y"]);
		expect(await loadedTabs(bytes)).toEqual(["X=1", "A=2", "Y=4"]);
		// Scope2 belonged to the dropped `a`; Y's Scope3 moves from raw index 3 to loaded index 2.
		expect((await openXlsx(bytes)).definedNames).toEqual([
			TOTAL,
			scope(0),
			scope(1),
			scope(3, 2),
		]);
		expect(await streamedA1(bytes, "Y")).toBe(4);
		await expect(streamedA1(bytes, "a")).rejects.toMatchObject({ code: "no-such-sheet" });
	});

	it("leaves out every later repeat, exact or case-only: [Data, Data, data]", async () => {
		const bytes = await tabsWorkbook(["Data", "Data", "data"]);
		expect(await loadedTabs(bytes)).toEqual(["Data=1"]);
		expect((await openXlsx(bytes)).definedNames).toEqual([TOTAL, scope(0)]);
	});

	it.each([
		[["A", "a"], ["A=1"]],
		[["_x0041_", "A"], ["A=1"]],
		[[KELVIN, "K"], [`${KELVIN}=1`]],
		[[DOTTED_I, I_DOT], [`${DOTTED_I}=1`]],
		[
			["X", "A", "a", "Y"],
			["X=1", "A=2", "Y=4"],
		],
	] as const)(
		"bridges only the kept tabs of %j, with their reconciled scoped names",
		async (names, kept) => {
			const before = await openXlsx(await tabsWorkbook(names));
			const bridged = await writeXlsx(await workbookToInput(before));
			expect(await loadedTabs(bridged)).toEqual(kept);
			expect((await openXlsx(bridged)).definedNames).toEqual(before.definedNames);
		},
	);
});

describe("reader and writer share one duplicate-name rule (F11.40(a))", () => {
	// Whether the reader leaves the later tab out, and whether writeXlsx refuses the same two names.
	// `toLowerCase` is JavaScript's full lower-case mapping: context-sensitive for final sigma, and
	// not Unicode case folding (sharp s and `SS` stay apart). Excel's own comparison is not known.
	const PAIRS: readonly (readonly [string, string, boolean])[] = [
		["A", "a", true],
		["A", "A ", false], // a trailing space is part of the name
		[KELVIN, "K", true],
		[DOTTED_I, I_DOT, true],
		[DOTTED_I, "i", false],
		["I", DOTLESS_I, false],
		["\u00DF", "SS", false], // sharp s
		["\u00DF", "ss", false],
		["\u1E9E", "\u00DF", true], // capital sharp s, sharp s
		["\u03A3", "\u03C3", true], // capital sigma, small sigma
		["\u03A3", "\u03C2", false], // capital sigma, final sigma: a lone capital lower-cases to small
		["A\u03A3", "a\u03C2", true], // after a letter, capital sigma lower-cases to final sigma
		["A\u03A3", "a\u03C3", false],
		["\uFF21", "\uFF41", true], // fullwidth A, fullwidth a
		["\uFF21", "A", false], // fullwidth A and ASCII A: no compatibility folding
		["\u01C4", "\u01C6", true], // DZ with caron, capital and small
		["\u01C5", "\u01C6", true], // its titlecase form and the small one
		["\u00E9", "e\u0301", false], // precomposed and decomposed e acute: no normalization
		["Data", "Data", true],
	];

	it.each(PAIRS)(
		"[%s, %s]: the later tab is left out = %s, exactly when the writer refuses",
		async (first, second, collides) => {
			const read = await openXlsx(await tabsWorkbook([first, second]));
			expect(read.sheets.map((s) => s.name)).toEqual(collides ? [first] : [first, second]);
			const written = writeXlsx({
				sheets: [
					{ name: first, rows: [[1]] },
					{ name: second, rows: [[2]] },
				],
			});
			if (collides) {
				await expect(written).rejects.toThrow(XlsxError);
				await expect(written).rejects.toMatchObject({
					code: "invalid-input",
					message: `duplicate sheet name "${second}" (case-insensitive)`,
				});
			} else {
				await expect(written).resolves.toBeInstanceOf(Uint8Array);
			}
		},
	);
});
