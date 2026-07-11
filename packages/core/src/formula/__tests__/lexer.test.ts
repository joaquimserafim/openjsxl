import { describe, expect, it } from "vitest";
import { FormulaError } from "../errors";
import { tokenize } from "../lexer";

// Compact "type:value" view of the token stream (minus the trailing eof), so a single string pins the
// tokenizer's decisions — the load-bearing ones being cell/name/function disambiguation and the
// `$`-prefixed partial references.
function toks(s: string): string {
	return tokenize(s)
		.filter((t) => t.type !== "eof")
		.map((t) => `${t.type}:${t.value}`)
		.join(" ");
}

describe("tokenize — cell / name / function disambiguation", () => {
	it("reads a bare cell as a cell", () => {
		expect(toks("A1")).toBe("cell:A1");
	});

	it("reads an absolute cell with its $ markers kept", () => {
		expect(toks("$A$1")).toBe("cell:$A$1");
		expect(toks("$A1")).toBe("cell:$A1");
		expect(toks("A$1")).toBe("cell:A$1");
	});

	it("treats a cell-shaped run followed by '(' as a function name, not a cell", () => {
		// LOG10 is a valid cell address on its own, but LOG10( is the LOG10 function.
		expect(toks("LOG10(A1)")).toBe("name:LOG10 op:( cell:A1 op:)");
		expect(toks("LOG10")).toBe("cell:LOG10");
	});

	it("treats a cell-shaped run followed by '!' as a sheet name, not a cell", () => {
		expect(toks("A1!B2")).toBe("name:A1 op:! cell:B2");
	});

	it("keeps a 4th letter or 8th digit out of a cell (falls back to a name)", () => {
		expect(toks("ABCD1")).toBe("name:ABCD1");
		expect(toks("A12345678")).toBe("name:A12345678");
	});

	it("reads $-prefixed partial references (whole-column / whole-row endpoints)", () => {
		expect(toks("$A:$B")).toBe("name:$A op:: name:$B");
		expect(toks("$2:$5")).toBe("name:$2 op:: name:$5");
		expect(toks("A:A")).toBe("name:A op:: name:A");
		expect(toks("1:1")).toBe("num:1 op:: num:1");
	});
});

describe("tokenize — literals & operators", () => {
	it("reads numbers incl. decimals, leading dot, and scientific notation", () => {
		expect(toks("3.14")).toBe("num:3.14");
		expect(toks(".5")).toBe("num:.5");
		expect(toks("1E+20")).toBe("num:1E+20");
		expect(toks("1e-3")).toBe("num:1e-3");
	});

	it("does not swallow a trailing e with no exponent digits into the number", () => {
		// `1E` is the number 1 followed by the name E (a defined name), not a malformed number.
		expect(toks("1+E")).toBe("num:1 op:+ name:E");
	});

	it("unescapes doubled quotes inside strings and sheet names", () => {
		expect(tokenize('"a""b"')[0]).toMatchObject({ type: "str", value: 'a"b' });
		expect(tokenize("'a''b'!A1")[0]).toMatchObject({ type: "sheet", value: "a'b" });
	});

	it("recognizes all eight cell-error literals as single tokens", () => {
		for (const e of [
			"#DIV/0!",
			"#N/A",
			"#NAME?",
			"#NULL!",
			"#NUM!",
			"#REF!",
			"#VALUE!",
			"#GETTING_DATA",
		]) {
			expect(toks(e)).toBe(`err:${e}`);
		}
	});

	it("reads a bare # (not an error literal) as the spill operator", () => {
		expect(toks("A1#")).toBe("cell:A1 op:#");
	});

	it("reads two-character comparison operators as one token", () => {
		expect(toks("1<=2")).toBe("num:1 op:<= num:2");
		expect(toks("1<>2")).toBe("num:1 op:<> num:2");
		expect(toks("1>=2")).toBe("num:1 op:>= num:2");
	});

	it("captures a balanced [...] span as one bracket token, nesting and all", () => {
		expect(toks("Table1[[#Data],[Amt]]")).toBe("name:Table1 bracket:[[#Data],[Amt]]");
	});

	it("drops insignificant whitespace", () => {
		expect(toks("  A1  +  B1 ")).toBe("cell:A1 op:+ cell:B1");
	});
});

describe("tokenize — typed failures", () => {
	it("throws a typed error on an unterminated string", () => {
		expect(() => tokenize('"abc')).toThrow(FormulaError);
	});

	it("throws a typed error on an unbalanced bracket", () => {
		expect(() => tokenize("Table1[abc")).toThrow(FormulaError);
	});
});
