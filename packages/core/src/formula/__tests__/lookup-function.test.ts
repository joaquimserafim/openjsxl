import { describe, expect, it } from "vitest";
import { buildRegistry, lookupFunction } from "../functions";

// F11.61: Excel stores a function added after Excel 2007 as `_xlfn.NAME`, and FILTER and SORT as
// `_xlfn._xlws.NAME`. lookupFunction resolves both to the registered NAME.

const builtins = buildRegistry(undefined);
const spec = (value: number) => ({ minArgs: 0, maxArgs: 2, evaluate: () => value });

describe("lookupFunction", () => {
	it("finds a name as written, in any case", () => {
		expect(builtins.get("SUM")).toBeDefined();
		expect(lookupFunction(builtins, "SUM")).toBe(builtins.get("SUM"));
		expect(lookupFunction(builtins, "sum")).toBe(builtins.get("SUM"));
		expect(lookupFunction(builtins, "Sum")).toBe(builtins.get("SUM"));
	});

	it("drops one leading _xlfn., in any case", () => {
		expect(builtins.get("IFNA")).toBeDefined();
		for (const name of ["_xlfn.IFNA", "_XLFN.IFNA", "_xlfn.ifna", "_XlFn.IfNa"]) {
			expect(lookupFunction(builtins, name)).toBe(builtins.get("IFNA"));
		}
	});

	it("drops an _xlws. that follows the _xlfn.", () => {
		expect(lookupFunction(builtins, "_xlfn._xlws.IFNA")).toBe(builtins.get("IFNA"));
		expect(lookupFunction(builtins, "_XLFN._XLWS.ifna")).toBe(builtins.get("IFNA"));
	});

	it("drops nothing else", () => {
		for (const name of [
			"NOSUCH",
			"_xlfn.NOSUCH",
			"_xlfn._xlfn.IFNA",
			"_xlfn._xlws._xlws.IFNA",
			"_xlws.IFNA",
			"_xlws._xlfn.IFNA",
			"_xlfn.",
			"_xlfn._xlws.",
			"xlfn.IFNA",
			"_xlfnIFNA",
			"_xlfn_IFNA",
			"_xlfn..IFNA",
			"IFNA_xlfn.",
			"_xludf.IFNA",
			"_xlpm.IFNA",
		]) {
			expect(lookupFunction(builtins, name)).toBeUndefined();
		}
	});

	it("drops _xlfn. only at the start of the name", () => {
		// Were it dropped anywhere, the first six characters would go and "xlfn.IFNA" would remain.
		const registry = buildRegistry({ "xlfn.IFNA": spec(6) });
		expect(registry.get("XLFN.IFNA")).toBeDefined();
		expect(lookupFunction(registry, "abcde_xlfn.IFNA")).toBeUndefined();
	});

	it("prefers a function registered under the prefixed name", () => {
		const registry = buildRegistry({ "_xlfn.IFNA": spec(1), "_xlfn._xlws.IFNA": spec(2) });
		const prefixed = registry.get("_XLFN.IFNA");
		const doubly = registry.get("_XLFN._XLWS.IFNA");
		expect(prefixed).toBeDefined();
		expect(doubly).toBeDefined();
		expect(lookupFunction(registry, "_xlfn.IFNA")).toBe(prefixed);
		expect(lookupFunction(registry, "_XLFN.ifna")).toBe(prefixed);
		expect(lookupFunction(registry, "_xlfn._xlws.IFNA")).toBe(doubly);
		expect(lookupFunction(registry, "IFNA")).toBe(builtins.get("IFNA"));
	});

	it("resolves a prefixed name to a caller's override of the bare name", () => {
		const registry = buildRegistry({ ifna: spec(3) });
		expect(registry.get("IFNA")).not.toBe(builtins.get("IFNA"));
		expect(lookupFunction(registry, "_xlfn.IFNA")).toBe(registry.get("IFNA"));
		expect(lookupFunction(registry, "_xlfn._xlws.IFNA")).toBe(registry.get("IFNA"));
	});

	it("folds only ASCII letters in the prefix", () => {
		// U+017F (long s) upper-cases to S, so full Unicode folding would read `_xlw\u017f.` as `_XLWS.`.
		for (const name of [
			"_xlfn._xlw\u017f.IFNA",
			"_XLFN._XLW\u017f.IFNA",
			"_xlfn._xlw\u017f.SUM",
		]) {
			expect(lookupFunction(builtins, name)).toBeUndefined();
		}
	});

	it("resolves the rest after the prefix exactly as the bare name", () => {
		// Including the Unicode folding bare names already get (U+0131 dotless i, the U+FB01 fi ligature).
		for (const name of ["IFNA", "ifna", "ERROR.TYPE", "\u0131fna", "\ufb01nd", "NOSUCH"]) {
			expect(lookupFunction(builtins, `_xlfn.${name}`)).toBe(lookupFunction(builtins, name));
		}
		expect(lookupFunction(builtins, "\u0131fna")).toBe(builtins.get("IFNA"));
	});

	it("never resolves a prefix alone, even to a function registered under the empty name", () => {
		const registry = buildRegistry({ "": spec(5) });
		expect(registry.get("")).toBeDefined();
		for (const name of ["_xlfn.", "_XLFN.", "_xlfn._xlws.", "_XLFN._XLWS."]) {
			expect(lookupFunction(registry, name)).toBeUndefined();
		}
	});

	it("never resolves a bare name to a prefixed registration", () => {
		const registry = buildRegistry({ "_xlfn.MYFN": spec(4) });
		expect(lookupFunction(registry, "MYFN")).toBeUndefined();
		expect(lookupFunction(registry, "_xlfn.MYFN")).toBe(registry.get("_XLFN.MYFN"));
	});
});
