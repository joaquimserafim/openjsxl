import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import type { Comment } from "../../types";
import { commentsParts } from "../comments";
import { definedNamesXml, validateDefinedNames } from "../defined-names";
import { sheetSideParts, validateSheetMeta, workbookProtectionXml } from "../parts";

// Run only synchronous production code while the prototype is changed. Restore it before any
// Vitest assertion, including failure reporting, so polluted fields cannot affect the runner.
function withInherited<T>(
	key: string,
	descriptor: PropertyDescriptor,
	run: () => T,
): { result?: T; error?: unknown } {
	const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
	try {
		Object.defineProperty(Object.prototype, key, { configurable: true, ...descriptor });
		try {
			return { result: run() };
		} catch (error) {
			return { error };
		}
	} finally {
		if (previous === undefined) Reflect.deleteProperty(Object.prototype, key);
		else Object.defineProperty(Object.prototype, key, previous);
	}
}

function emptySlot<T>(): T[] {
	const values: T[] = [];
	values.length = 1;
	return values;
}

describe("writer metadata ignores inherited model fields", () => {
	it.each([
		{ key: "name", value: "Injected", run: () => validateDefinedNames([{ refersTo: "1" }], 1) },
		{ key: "refersTo", value: "99", run: () => validateDefinedNames([{ name: "Own" }], 1) },
		{
			key: "ref",
			value: "B2",
			run: () => {
				// @ts-expect-error A required field is deliberately inherited rather than owned.
				return commentsParts("S", [{ text: "Own" }]);
			},
		},
		{
			key: "text",
			value: "Injected",
			run: () => {
				// @ts-expect-error A required field is deliberately inherited rather than owned.
				return commentsParts("S", [{ ref: "A1" }]);
			},
		},
		{ key: "name", value: "Injected", run: () => validateSheetMeta([{}]) },
	])("rejects inherited-only required $key", ({ key, value, run }) => {
		const outcome = withInherited<unknown>(key, { value, writable: true }, run);
		expect(outcome.error).toBeInstanceOf(XlsxError);
		expect(outcome.error).toMatchObject({ code: "invalid-input" });
	});

	it.each(["hidden", "localSheetId"])(
		"omits inherited defined-name %s after normalization",
		(key) => {
			const outcome = withInherited(
				key,
				{ value: key === "hidden" ? true : 0, writable: true },
				() =>
					definedNamesXml(
						validateDefinedNames([{ name: "Own", refersTo: '"_x0041_<&"' }], 1),
					),
			);
			expect(outcome.error).toBeUndefined();
			expect(outcome.result).toBe(
				'<definedNames><definedName name="Own">"_x005F_x0041_&lt;&amp;"</definedName></definedNames>',
			);
		},
	);

	it("does not resurrect an explicitly false hidden flag on a normalized name", () => {
		const outcome = withInherited("hidden", { value: true, writable: true }, () =>
			definedNamesXml(
				validateDefinedNames([{ name: "Own", refersTo: "1", hidden: false }], 1),
			),
		);
		expect(outcome.error).toBeUndefined();
		expect(outcome.result).toBe(
			'<definedNames><definedName name="Own">1</definedName></definedNames>',
		);
	});

	it("does not evaluate an inherited author getter or change anonymous author deduplication", () => {
		let reads = 0;
		const outcome = withInherited(
			"author",
			{
				get: () => {
					reads++;
					throw new Error("inherited getter");
				},
			},
			() =>
				commentsParts("S", [
					{ ref: "A1", text: "_x0041_\r" },
					{ ref: "B2", text: "Own" },
				]),
		);
		expect(outcome.error).toBeUndefined();
		expect(reads).toBe(0);
		expect(outcome.result?.commentsXml).toContain("<authors><author></author></authors>");
		expect(outcome.result?.commentsXml).toContain("_x005F_x0041__x000D_");
		expect(outcome.result?.commentsXml.match(/authorId="0"/g)).toHaveLength(2);
	});

	it("does not evaluate an inherited sheet-state getter", () => {
		let reads = 0;
		const outcome = withInherited(
			"state",
			{
				get: () => {
					reads++;
					throw new Error("inherited getter");
				},
			},
			() => validateSheetMeta([{ name: "Own" }]),
		);
		expect(outcome.error).toBeUndefined();
		expect(reads).toBe(0);
		expect(outcome.result).toEqual({ names: ["Own"], states: ["visible"] });
	});

	it.each([
		"lockStructure",
		"lockWindows",
		"workbookPassword",
		"workbookAlgorithmName",
		"workbookHashValue",
		"workbookSaltValue",
		"workbookSpinCount",
	])("does not evaluate inherited workbook protection %s", (key) => {
		let reads = 0;
		const outcome = withInherited(
			key,
			{
				get: () => {
					reads++;
					throw new Error("inherited getter");
				},
			},
			() => workbookProtectionXml({}),
		);
		expect(outcome.error).toBeUndefined();
		expect(outcome.result).toBe("");
		expect(reads).toBe(0);
	});

	it.each(["commentsXml", "vmlXml", "drawingXml", "drawingRelsXml", "tables"])(
		"does not evaluate inherited optional side-part %s",
		(key) => {
			let reads = 0;
			const outcome = withInherited(
				key,
				{
					get: () => {
						reads++;
						throw new Error("inherited getter");
					},
				},
				() => sheetSideParts(0, { rels: [] }),
			);
			expect(outcome.error).toBeUndefined();
			expect(outcome.result).toEqual([]);
			expect(reads).toBe(0);
		},
	);

	it.each([
		{
			label: "defined names",
			value: { name: "Injected", refersTo: "1" },
			run: () => validateDefinedNames(emptySlot<unknown>(), 1),
		},
		{
			label: "comments",
			value: { ref: "A1", text: "Injected" },
			run: () => commentsParts("S", emptySlot<Comment>()),
		},
		{
			label: "sheet metadata",
			value: { name: "Injected" },
			run: () => validateSheetMeta(emptySlot<unknown>()),
		},
	])("rejects an inherited array slot in $label", ({ run, value }) => {
		const outcome = withInherited<unknown>("0", { value, writable: true }, run);
		expect(outcome.error).toBeInstanceOf(XlsxError);
		expect(outcome.error).toMatchObject({ code: "invalid-input" });
	});
});

describe("writer metadata retains own-property behavior", () => {
	it("accepts null-prototype names, comments, sheet metadata and protection", () => {
		const name = Object.setPrototypeOf(
			{ name: "Own", refersTo: "1", localSheetId: 0, hidden: true },
			null,
		);
		const comment = Object.setPrototypeOf(
			{ ref: "B2", text: "  _x0041_\r  ", author: "_x0042_" },
			null,
		);
		const sheet = Object.setPrototypeOf({ name: "Own", state: "visible" }, null);
		const protection = Object.setPrototypeOf(
			{ lockStructure: true, workbookPassword: "a<&\r", workbookSpinCount: 0 },
			null,
		);
		expect(validateDefinedNames([name], 1)).toEqual([name]);
		expect(commentsParts("S", [comment])).toEqual(commentsParts("S", [{ ...comment }]));
		expect(validateSheetMeta([sheet])).toEqual({ names: ["Own"], states: ["visible"] });
		expect(workbookProtectionXml(protection)).toBe(
			'<workbookProtection lockStructure="1" workbookPassword="a&lt;&amp;&#13;" workbookSpinCount="0"/>',
		);
	});

	it("reads own getters once, including non-enumerable fields", () => {
		const reads = new Map<string, number>();
		function field(key: string, value: unknown): PropertyDescriptor {
			return {
				get: () => {
					reads.set(key, (reads.get(key) ?? 0) + 1);
					return value;
				},
			};
		}
		const name = Object.defineProperties(
			{},
			{
				name: field("name", "Own"),
				refersTo: field("refersTo", "1"),
				hidden: field("hidden", false),
			},
		);
		const comment = Object.defineProperties(
			{ ref: "A1", text: "Note" },
			{
				ref: field("ref", "B2"),
				text: field("text", "_x0041_"),
				author: field("author", "_x0042_"),
			},
		);
		const sheet = Object.defineProperties(
			{},
			{ name: field("sheetName", "Own"), state: field("state", "visible") },
		);
		const protection = Object.defineProperties(
			{},
			{
				lockStructure: field("lockStructure", true),
				workbookPassword: field("workbookPassword", "_x0041_"),
				workbookSpinCount: field("workbookSpinCount", 0),
			},
		);
		expect(definedNamesXml(validateDefinedNames([name], 1))).not.toContain("hidden=");
		expect(commentsParts("S", [comment])?.commentsXml).toContain('<comment ref="B2"');
		expect(validateSheetMeta([sheet]).names).toEqual(["Own"]);
		expect(workbookProtectionXml(protection)).toContain('lockStructure="1"');
		expect([...reads.values()]).toEqual(Array(11).fill(1));
	});

	it.each([
		{
			label: "defined name",
			run: (read: () => void) =>
				validateDefinedNames(
					[
						{
							get name() {
								read();
								return "A1";
							},
							refersTo: "1",
						},
					],
					1,
				),
		},
		{
			label: "comment reference",
			run: (read: () => void) =>
				commentsParts("S", [
					{
						get ref() {
							read();
							return "A0";
						},
						text: "Own",
					},
				]),
		},
		{
			label: "sheet state",
			run: (read: () => void) =>
				validateSheetMeta([
					{
						name: "Own",
						get state() {
							read();
							return "invalid";
						},
					},
				]),
		},
		{
			label: "workbook protection",
			run: (read: () => void) =>
				workbookProtectionXml({
					get workbookSpinCount() {
						read();
						return -1;
					},
				}),
		},
	])("reads the invalid own $label getter only once", ({ run }) => {
		let reads = 0;
		let error: unknown;
		try {
			run(() => {
				reads++;
			});
		} catch (cause) {
			error = cause;
		}
		expect(error).toBeInstanceOf(XlsxError);
		expect(error).toMatchObject({ code: "invalid-input" });
		expect(reads).toBe(1);
	});

	it.each([
		{
			label: "defined names",
			run: () => validateDefinedNames([{ name: "Own", refersTo: "1", unknownField: 1 }], 1),
		},
		{
			label: "comments",
			run: () => commentsParts("S", [{ ref: "A1", text: "Own", ...{ unknownField: 1 } }]),
		},
		{
			label: "workbook protection",
			run: () => workbookProtectionXml({ ...{ unknownField: 1 }, lockStructure: false }),
		},
	])("continues rejecting unknown own keys in $label", ({ run }) => {
		expect(run).toThrowError(XlsxError);
		expect(run).toThrowError(/unknown property/);
	});

	it("ignores unknown inherited properties", () => {
		let reads = 0;
		const outcome = withInherited(
			"unknownField",
			{
				enumerable: true,
				get: () => {
					reads++;
					throw new Error("inherited getter");
				},
			},
			() => {
				validateDefinedNames([{ name: "Own", refersTo: "1" }], 1);
				commentsParts("S", [{ ref: "A1", text: "Own" }]);
				return workbookProtectionXml({});
			},
		);
		expect(outcome.error).toBeUndefined();
		expect(outcome.result).toBe("");
		expect(reads).toBe(0);
	});
});
