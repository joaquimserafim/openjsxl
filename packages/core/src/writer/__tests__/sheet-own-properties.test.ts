import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { createMediaRegistry } from "../images";
import { createTableContext, worksheetXml } from "../sheet";
import { createStyleRegistry } from "../styles";
import type { SheetInput } from "../types";

function inherited<T>(key: string, value: unknown, run: () => T): T {
	const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
	Object.defineProperty(Object.prototype, key, { configurable: true, writable: true, value });
	try {
		return run();
	} finally {
		if (previous === undefined) Reflect.deleteProperty(Object.prototype, key);
		else Object.defineProperty(Object.prototype, key, previous);
	}
}

function feature(key: string, value: unknown): SheetInput {
	const sheet: SheetInput = { name: "Data", rows: [["Header"], [2]] };
	Object.defineProperty(sheet, key, { enumerable: true, configurable: true, value });
	return sheet;
}

function render(sheet: SheetInput) {
	const styles = createStyleRegistry();
	const result = worksheetXml(
		sheet,
		sheet.rows,
		0,
		false,
		styles,
		createMediaRegistry(),
		createTableContext(),
		"Data",
	);
	return { result, styles: styles.stylesXml() };
}

type Case = {
	label: string;
	key: string;
	inherited: unknown;
	record: () => object;
	wrap: (record: object) => SheetInput;
};
const cases: Case[] = [
	{
		label: "image extent",
		key: "cx",
		inherited: 10,
		record: () => ({ cy: 1 }),
		wrap: (r) =>
			feature("images", [
				{
					bytes: new Uint8Array([1]),
					mime: "image/png",
					anchor: { from: { col: 1, row: 1 }, ext: r },
				},
			]),
	},
	{
		label: "column",
		key: "width",
		inherited: 20,
		record: () => ({ min: 1, max: 1 }),
		wrap: (r) => feature("columns", [r]),
	},
	{
		label: "row",
		key: "height",
		inherited: 40,
		record: () => ({}),
		wrap: (r) => feature("rowProperties", { 1: r }),
	},
	{
		label: "freeze",
		key: "cols",
		inherited: 2,
		record: () => ({ rows: 1 }),
		wrap: (r) => feature("freeze", r),
	},
	{
		label: "protection",
		key: "sheet",
		inherited: true,
		record: () => ({}),
		wrap: (r) => feature("protection", r),
	},
	{
		label: "print options",
		key: "headings",
		inherited: true,
		record: () => ({}),
		wrap: (r) => feature("printOptions", r),
	},
	{
		label: "page setup",
		key: "orientation",
		inherited: "landscape",
		record: () => ({}),
		wrap: (r) => feature("pageSetup", r),
	},
	{
		label: "header/footer",
		key: "oddHeader",
		inherited: "injected",
		record: () => ({}),
		wrap: (r) => feature("headerFooter", r),
	},
	{
		label: "hyperlink",
		key: "tooltip",
		inherited: "injected",
		record: () => ({ ref: "A1", location: "Data!A1" }),
		wrap: (r) => feature("hyperlinks", [r]),
	},
	{
		label: "data validation",
		key: "prompt",
		inherited: "injected",
		record: () => ({ type: "whole", sqref: ["A1"], formula1: "1" }),
		wrap: (r) => feature("dataValidations", [r]),
	},
	{
		label: "CF rule",
		key: "stopIfTrue",
		inherited: true,
		record: () => ({ type: "expression", priority: 1, formulas: ["1"] }),
		wrap: (r) => feature("conditionalFormatting", [{ sqref: ["A1"], rules: [r] }]),
	},
	{
		label: "CF threshold",
		key: "gte",
		inherited: false,
		record: () => ({ type: "min" }),
		wrap: (r) =>
			feature("conditionalFormatting", [
				{
					sqref: ["A1"],
					rules: [
						{
							type: "dataBar",
							priority: 1,
							cfvo: [r, { type: "max" }],
							color: { rgb: "112233" },
						},
					],
				},
			]),
	},
	{
		label: "table",
		key: "totalsRow",
		inherited: true,
		record: () => ({ name: "Sales", ref: "A1:A2" }),
		wrap: (r) => feature("tables", [r]),
	},
	{
		label: "table style",
		key: "showRowStripes",
		inherited: true,
		record: () => ({}),
		wrap: (r) => feature("tables", [{ name: "Sales", ref: "A1:A2", style: r }]),
	},
	{
		label: "table column",
		key: "totalsRowLabel",
		inherited: "injected",
		record: () => ({ name: "Header" }),
		wrap: (r) => feature("tables", [{ name: "Sales", ref: "A1:A2", columns: [r] }]),
	},
	{
		label: "styled cell",
		key: "formula",
		inherited: "1+1",
		record: () => ({ value: 4 }),
		wrap: (r) => feature("rows", [[r]]),
	},
	{
		label: "image",
		key: "name",
		inherited: "injected",
		record: () => ({
			bytes: new Uint8Array([1]),
			mime: "image/png",
			anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
		}),
		wrap: (r) => feature("images", [r]),
	},
	{
		label: "image anchor",
		key: "editAs",
		inherited: "absolute",
		record: () => ({ from: { col: 1, row: 1 }, to: { col: 2, row: 2 } }),
		wrap: (r) =>
			feature("images", [{ bytes: new Uint8Array([1]), mime: "image/png", anchor: r }]),
	},
	{
		label: "image point",
		key: "colOff",
		inherited: 10,
		record: () => ({ col: 1, row: 1 }),
		wrap: (r) =>
			feature("images", [
				{
					bytes: new Uint8Array([1]),
					mime: "image/png",
					anchor: { from: r, ext: { cx: 1, cy: 1 } },
				},
			]),
	},
];

describe("worksheet own model properties", () => {
	it.each(cases)(
		"ignores inherited optional $label fields",
		({ key, inherited: value, record, wrap }) => {
			const expected = render(wrap(record()));
			const actual = inherited(key, value, () => render(wrap(record())));
			expect(actual).toEqual(expected);
		},
	);
	it.each(cases)(
		"retains own and null-prototype $label fields, reading getters once",
		({ key, inherited: value, record, wrap }) => {
			const own = record();
			Object.defineProperty(own, key, { enumerable: true, value });
			const expected = render(wrap(own));
			const plain = record();
			Object.setPrototypeOf(plain, null);
			let reads = 0;
			Object.defineProperty(plain, key, {
				enumerable: true,
				get() {
					reads++;
					return value;
				},
			});
			const actual = render(wrap(plain));
			expect(actual).toEqual(expected);
			expect(reads).toBe(1);
		},
	);
	it.each(cases)(
		"rejects unknown own $label fields without reading values twice",
		({ key, record, wrap }) => {
			const raw = record();
			let reads = 0;
			Object.defineProperty(raw, key, {
				enumerable: true,
				get() {
					reads++;
					return undefined;
				},
			});
			Object.defineProperty(raw, "unexpected", { enumerable: true, value: true });
			expect(() => render(wrap(raw))).toThrow(XlsxError);
			expect(reads).toBeLessThanOrEqual(1);
		},
	);
	it("requires an own value or formula on object cells", () => {
		expect(() => inherited("value", 7, () => render(feature("rows", [[{}]])))).toThrow(
			XlsxError,
		);
		expect(() => render(feature("rows", [[Object.create({ value: 7 })]]))).toThrow(XlsxError);
		// Presence, rather than the resolved value, selects the object-cell shape.
		expect(() => render(feature("rows", [[{ value: undefined }]]))).not.toThrow();
	});
});

const requiredCases: Case[] = [
	{
		label: "column min",
		key: "min",
		inherited: 1,
		record: () => ({ max: 1 }),
		wrap: (r) => feature("columns", [r]),
	},
	{
		label: "column max",
		key: "max",
		inherited: 1,
		record: () => ({ min: 1 }),
		wrap: (r) => feature("columns", [r]),
	},
	{
		label: "filter ref",
		key: "ref",
		inherited: "A1:A2",
		record: () => ({}),
		wrap: (r) => feature("autoFilter", r),
	},
	{
		label: "margin",
		key: "left",
		inherited: 1,
		record: () => ({ right: 1, top: 1, bottom: 1, header: 1, footer: 1 }),
		wrap: (r) => feature("pageMargins", r),
	},
	{
		label: "hyperlink ref",
		key: "ref",
		inherited: "A1",
		record: () => ({ location: "Data!A1" }),
		wrap: (r) => feature("hyperlinks", [r]),
	},
	{
		label: "hyperlink destination",
		key: "location",
		inherited: "Data!A1",
		record: () => ({ ref: "A1" }),
		wrap: (r) => feature("hyperlinks", [r]),
	},
	{
		label: "validation sqref",
		key: "sqref",
		inherited: ["A1"],
		record: () => ({}),
		wrap: (r) => feature("dataValidations", [r]),
	},
	{
		label: "CF sqref",
		key: "sqref",
		inherited: ["A1"],
		record: () => ({ rules: [{ type: "expression", priority: 1 }] }),
		wrap: (r) => feature("conditionalFormatting", [r]),
	},
	{
		label: "CF rules",
		key: "rules",
		inherited: [{ type: "expression", priority: 1 }],
		record: () => ({ sqref: ["A1"] }),
		wrap: (r) => feature("conditionalFormatting", [r]),
	},
	{
		label: "CF type",
		key: "type",
		inherited: "expression",
		record: () => ({ priority: 1 }),
		wrap: (r) => feature("conditionalFormatting", [{ sqref: ["A1"], rules: [r] }]),
	},
	{
		label: "CF priority",
		key: "priority",
		inherited: 1,
		record: () => ({ type: "expression" }),
		wrap: (r) => feature("conditionalFormatting", [{ sqref: ["A1"], rules: [r] }]),
	},
	{
		label: "CF threshold type",
		key: "type",
		inherited: "min",
		record: () => ({}),
		wrap: (r) =>
			feature("conditionalFormatting", [
				{
					sqref: ["A1"],
					rules: [
						{
							type: "dataBar",
							priority: 1,
							cfvo: [r, { type: "max" }],
							color: { rgb: "112233" },
						},
					],
				},
			]),
	},
	{
		label: "table name",
		key: "name",
		inherited: "Sales",
		record: () => ({ ref: "A1:A2" }),
		wrap: (r) => feature("tables", [r]),
	},
	{
		label: "table ref",
		key: "ref",
		inherited: "A1:A2",
		record: () => ({ name: "Sales" }),
		wrap: (r) => feature("tables", [r]),
	},
	{
		label: "image bytes",
		key: "bytes",
		inherited: new Uint8Array([1]),
		record: () => ({
			mime: "image/png",
			anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
		}),
		wrap: (r) => feature("images", [r]),
	},
	{
		label: "image mime",
		key: "mime",
		inherited: "image/png",
		record: () => ({
			bytes: new Uint8Array([1]),
			anchor: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
		}),
		wrap: (r) => feature("images", [r]),
	},
	{
		label: "image anchor",
		key: "anchor",
		inherited: { from: { col: 1, row: 1 }, ext: { cx: 1, cy: 1 } },
		record: () => ({ bytes: new Uint8Array([1]), mime: "image/png" }),
		wrap: (r) => feature("images", [r]),
	},
	{
		label: "image from",
		key: "from",
		inherited: { col: 1, row: 1 },
		record: () => ({ ext: { cx: 1, cy: 1 } }),
		wrap: (r) =>
			feature("images", [{ bytes: new Uint8Array([1]), mime: "image/png", anchor: r }]),
	},
	{
		label: "point col",
		key: "col",
		inherited: 1,
		record: () => ({ row: 1 }),
		wrap: (r) =>
			feature("images", [
				{
					bytes: new Uint8Array([1]),
					mime: "image/png",
					anchor: { from: r, ext: { cx: 1, cy: 1 } },
				},
			]),
	},
	{
		label: "point row",
		key: "row",
		inherited: 1,
		record: () => ({ col: 1 }),
		wrap: (r) =>
			feature("images", [
				{
					bytes: new Uint8Array([1]),
					mime: "image/png",
					anchor: { from: r, ext: { cx: 1, cy: 1 } },
				},
			]),
	},
];

describe("required own worksheet fields", () => {
	it.each(requiredCases)(
		"rejects inherited $label",
		({ key, inherited: value, record, wrap }) => {
			expect(() => inherited(key, value, () => render(wrap(record())))).toThrow(XlsxError);
		},
	);
	it.each(requiredCases)(
		"accepts own null-prototype $label once",
		({ key, inherited: value, record, wrap }) => {
			const raw = record();
			Object.setPrototypeOf(raw, null);
			let reads = 0;
			Object.defineProperty(raw, key, {
				enumerable: true,
				get() {
					reads++;
					return value;
				},
			});
			expect(() => render(wrap(raw))).not.toThrow();
			expect(reads).toBe(1);
		},
	);
	it.each(requiredCases)("rejects invalid own $label once", ({ key, record, wrap }) => {
		const raw = record();
		let reads = 0;
		Object.defineProperty(raw, key, {
			enumerable: true,
			get() {
				reads++;
				return undefined;
			},
		});
		expect(() => render(wrap(raw))).toThrow(XlsxError);
		expect(reads).toBe(1);
	});
});
