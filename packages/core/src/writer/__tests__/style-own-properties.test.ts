import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import type { CellStyle } from "../../types";
import { colorXml, createStyleRegistry, validateColor } from "../styles";

// Keep pollution synchronous and restore it before assertion/library code runs.
function inherited<T>(key: string, descriptor: PropertyDescriptor, run: () => T): T {
	const previous = Object.getOwnPropertyDescriptor(Object.prototype, key);
	Object.defineProperty(Object.prototype, key, { configurable: true, ...descriptor });
	try {
		return run();
	} finally {
		if (previous === undefined) Reflect.deleteProperty(Object.prototype, key);
		else Object.defineProperty(Object.prototype, key, previous);
	}
}

function cellResult(style: CellStyle, isDate = false) {
	const registry = createStyleRegistry();
	const index = registry.xfIndexFor(style, isDate, "A1");
	return { index, xml: registry.stylesXml(), theme: registry.usesTheme() };
}

function dxfResult(dxf: unknown) {
	const registry = createStyleRegistry();
	const index = registry.internDxf(dxf, "rule 1");
	return { index, xml: registry.stylesXml(), theme: registry.usesTheme() };
}

const fail = (message: string): never => {
	throw new XlsxError("invalid-input", message);
};

function styleAt(path: readonly string[], leaf: object): CellStyle {
	const style: CellStyle = {};
	let parent: object = style;
	for (const [index, key] of path.entries()) {
		const value = index === path.length - 1 ? leaf : {};
		Object.defineProperty(parent, key, { enumerable: true, value });
		parent = value;
	}
	return style;
}

const consumers: readonly {
	name: string;
	base: () => Record<string, unknown>;
	key: string;
	valid: unknown;
	invalid: unknown;
	read: (raw: Record<string, unknown>) => unknown;
}[] = [
	{
		name: "cell style",
		base: () => ({}),
		key: "font",
		valid: { bold: true },
		invalid: false,
		read: (raw) => cellResult(raw),
	},
	{
		name: "font",
		base: () => ({}),
		key: "name",
		valid: "Arial",
		invalid: 0,
		read: (raw) => cellResult(styleAt(["font"], raw)),
	},
	{
		name: "cell fill",
		base: () => ({ patternType: "solid" }),
		key: "fgColor",
		valid: { rgb: "112233" },
		invalid: {},
		read: (raw) => cellResult(styleAt(["fill"], raw)),
	},
	{
		name: "border",
		base: () => ({}),
		key: "top",
		valid: { style: "thin" },
		invalid: {},
		read: (raw) => cellResult(styleAt(["border"], raw)),
	},
	{
		name: "border edge",
		base: () => ({ style: "thin" }),
		key: "color",
		valid: { rgb: "112233" },
		invalid: {},
		read: (raw) => cellResult(styleAt(["border", "top"], raw)),
	},
	{
		name: "alignment",
		base: () => ({}),
		key: "horizontal",
		valid: "center",
		invalid: "invalid",
		read: (raw) => cellResult(styleAt(["alignment"], raw)),
	},
	{
		name: "protection",
		base: () => ({}),
		key: "locked",
		valid: false,
		invalid: 0,
		read: (raw) => cellResult(styleAt(["protection"], raw)),
	},
	{
		name: "differential style",
		base: () => ({}),
		key: "numberFormat",
		valid: "0.000",
		invalid: 0,
		read: dxfResult,
	},
	{
		name: "differential fill",
		base: () => ({}),
		key: "bgColor",
		valid: { rgb: "112233" },
		invalid: {},
		read: (raw) => dxfResult({ fill: raw }),
	},
	{
		name: "color",
		base: () => ({ theme: 4 }),
		key: "tint",
		valid: 0.5,
		invalid: "invalid",
		read: (raw) => colorXml("color", validateColor(fail, "color", raw)),
	},
];

for (const { name, base, key, valid, invalid, read } of consumers) {
	describe(`${name} own-property boundary`, () => {
		it("ignores inherited optional values without reading getters", () => {
			let reads = 0;
			const expected = read(base());
			const actual = inherited(
				key,
				{
					get() {
						reads++;
						return valid;
					},
				},
				() => read(base()),
			);
			expect(reads).toBe(0);
			expect(actual).toEqual(expected);
		});

		it("accepts own values and null-prototype records with identical output", () => {
			const ordinary = { ...base(), [key]: valid };
			const nullPrototype = { ...base(), [key]: valid };
			Object.setPrototypeOf(nullPrototype, null);
			expect(read(nullPrototype)).toEqual(read(ordinary));
		});

		it("reads own getters once before emission", () => {
			let reads = 0;
			const raw = base();
			Object.defineProperty(raw, key, {
				enumerable: true,
				get() {
					if (++reads > 1) throw new Error("second read");
					return valid;
				},
			});
			expect(read(raw)).toEqual(read({ ...base(), [key]: valid }));
			expect(reads).toBe(1);
		});

		it("rejects the first invalid getter result with a typed error and one read", () => {
			let reads = 0;
			const raw = base();
			Object.defineProperty(raw, key, {
				enumerable: true,
				get() {
					return ++reads === 1 ? invalid : valid;
				},
			});
			expect(() => read(raw)).toThrow(expect.objectContaining({ code: "invalid-input" }));
			expect(reads).toBe(1);
		});

		it("keeps unknown enumerable own keys invalid and ignores other unknown keys", () => {
			expect(() => read({ ...base(), unexpected: true })).toThrow(/unknown property/);
			const raw = base();
			Object.defineProperty(raw, "unexpected", { value: true });
			Object.defineProperty(raw, Symbol("unexpected"), { value: true, enumerable: true });
			const expected = read(base());
			expect(read(raw)).toEqual(expected);
			expect(
				inherited("unexpected", { value: true, enumerable: true }, () => read(base())),
			).toEqual(expected);
		});

		it("keeps known non-enumerable own fields readable", () => {
			const raw = base();
			Object.defineProperty(raw, key, { value: valid });
			expect(read(raw)).toEqual(read({ ...base(), [key]: valid }));
		});
	});
}

describe("style consumers use own properties", () => {
	it("ignores inherited top-level style components and number formats", () => {
		const expected = cellResult({}, true);
		const actual = inherited("numberFormat", { value: "0.00", writable: true }, () =>
			cellResult({}, true),
		);
		expect(actual).toEqual(expected);
	});

	it("does not inherit optional font fields into normalized emission", () => {
		const expected = cellResult({ font: { italic: true } });
		const actual = inherited("bold", { value: true, writable: true }, () =>
			cellResult({ font: { italic: true } }),
		);
		expect(actual).toEqual(expected);
	});

	it("does not accept an inherited required fill pattern", () => {
		const style: CellStyle = {};
		Object.defineProperty(style, "fill", { enumerable: true, value: {} });
		expect(() =>
			inherited("patternType", { value: "solid", writable: true }, () => cellResult(style)),
		).toThrow(XlsxError);
	});

	it("does not let inherited color variants shadow an own variant", () => {
		const expected = cellResult({ font: { color: { theme: 2 } } });
		const actual = inherited("rgb", { value: "AABBCC", writable: true }, () =>
			cellResult({ font: { color: { theme: 2 } } }),
		);
		expect(actual).toEqual(expected);
	});

	it("does not call inherited getters while resolving differential styles", () => {
		let reads = 0;
		const expected = dxfResult({ fill: { bgColor: { rgb: "112233" } } });
		const actual = inherited(
			"font",
			{
				get() {
					reads++;
					return { bold: true };
				},
			},
			() => dxfResult({ fill: { bgColor: { rgb: "112233" } } }),
		);
		expect(reads).toBe(0);
		expect(actual).toEqual(expected);
	});

	it("retains own undefined color discriminants as invalid input", () => {
		expect(() => validateColor(fail, "color", { rgb: undefined })).toThrow(/color.rgb/);
	});

	it.each(["rgb", "theme", "indexed", "auto"])("requires an own %s color discriminant", (key) => {
		const value = key === "rgb" ? "112233" : key === "auto" ? true : 2;
		expect(() => inherited(key, { value }, () => validateColor(fail, "color", {}))).toThrow(
			expect.objectContaining({ code: "invalid-input" }),
		);
	});

	it("requires an own border line style", () => {
		expect(() =>
			inherited("style", { value: "thin" }, () => cellResult(styleAt(["border", "top"], {}))),
		).toThrow(expect.objectContaining({ code: "invalid-input" }));
	});

	it("does not call inherited setters while constructing normalized components", () => {
		let writes = 0;
		const expected = cellResult({ font: { bold: true } });
		const actual = inherited(
			"bold",
			{
				set() {
					writes++;
				},
			},
			() => cellResult({ font: { bold: true } }),
		);
		expect(writes).toBe(0);
		expect(actual).toEqual(expected);
	});

	it("does not use inherited toJSON to collapse distinct interned styles", () => {
		const run = () => {
			const registry = createStyleRegistry();
			const first = registry.xfIndexFor({ font: { bold: true } }, false, "A1");
			const second = registry.xfIndexFor({ font: { italic: true } }, false, "B1");
			const dxfA = registry.internDxf({ fill: { bgColor: { rgb: "112233" } } }, "rule 1");
			const dxfB = registry.internDxf({ fill: { bgColor: { rgb: "445566" } } }, "rule 2");
			return { first, second, dxfA, dxfB, xml: registry.stylesXml() };
		};
		const expected = run();
		let calls = 0;
		const actual = inherited(
			"toJSON",
			{
				value() {
					calls++;
					return "collapsed";
				},
			},
			run,
		);
		expect(calls).toBe(0);
		expect(actual).toEqual(expected);
	});
});
