import { describe, expect, it } from "vitest";
import type { FormulaAst, SheetSpec } from "../ast";
import { FormulaError, type FormulaErrorCode } from "../errors";
import { parseFormula } from "../parser";

// A test-only reprinter that FULLY parenthesizes every operator, so one expected string pins both
// precedence and associativity unambiguously. The production module intentionally ships no
// serializer (F8.1 scope-out) — this lives here, in the test, only to make the tree legible.
function sheet(s?: SheetSpec): string {
	if (s === undefined) return "";
	return `${s.name}${s.toName !== undefined ? `:${s.toName}` : ""}!`;
}
function pp(n: FormulaAst): string {
	switch (n.type) {
		case "number":
			return n.raw;
		case "string":
			return JSON.stringify(n.value);
		case "boolean":
			return n.value ? "TRUE" : "FALSE";
		case "error":
			return n.value;
		case "cell":
			return `${sheet(n.sheet)}${n.ref}`;
		case "name":
			return `${sheet(n.sheet)}${n.name}`;
		case "range":
			return `(${pp(n.left)}:${pp(n.right)})`;
		case "union":
			return `(${n.items.map(pp).join(",")})`;
		case "binary":
			return `(${pp(n.left)}${n.op}${pp(n.right)})`;
		case "unary":
			return `(${n.op}${pp(n.operand)})`;
		case "postfix":
			return `(${pp(n.operand)}${n.op})`;
		case "group":
			return `[${pp(n.expr)}]`;
		case "call":
			return `${n.name}(${n.args.map(pp).join(",")})`;
		case "empty":
			return "";
		case "array":
			return `{${n.rows.map((r) => r.map(pp).join(",")).join(";")}}`;
		case "structured":
		case "external":
			return n.source;
	}
}
const rp = (f: string) => pp(parseFormula(f));

describe("parseFormula — precedence & associativity", () => {
	// Each row is Excel's documented result (MS operator-precedence table + M8 decision 6).
	it.each([
		["1+2*3", "(1+(2*3))"],
		["2*3+4", "((2*3)+4)"],
		["-2^2", "((-2)^2)"], // unary minus binds ABOVE ^  → (-2)^2 = 4
		["2^3^2", "((2^3)^2)"], // ^ is left-associative → 64
		["2^-2", "(2^(-2))"],
		["-2%", "((-2)%)"], // unary minus binds above %
		["1<2=3", "((1<2)=3)"], // comparisons left-associative
		["1&2&3", "((1&2)&3)"], // & left-associative
		["A1&B1=C1", "((A1&B1)=C1)"], // & (tighter) before comparison
		["10/2/5", "((10/2)/5)"],
		["-A1^2", "((-A1)^2)"],
		["2*-3", "(2*(-3))"],
		["--1", "(-(-1))"],
		["1+2-3", "((1+2)-3)"],
	])("%s → %s", (input, expected) => {
		expect(rp(input)).toBe(expected);
	});
});

describe("parseFormula — literals", () => {
	it("parses number forms", () => {
		expect(parseFormula("3")).toMatchObject({ type: "number", value: 3 });
		expect(parseFormula("3.14")).toMatchObject({ type: "number", value: 3.14 });
		expect(parseFormula(".5")).toMatchObject({ type: "number", value: 0.5 });
		expect(parseFormula("1E+20")).toMatchObject({ type: "number", value: 1e20, raw: "1E+20" });
	});

	it("parses strings and unescapes doubled quotes", () => {
		expect(parseFormula('"hi"')).toMatchObject({ type: "string", value: "hi" });
		expect(parseFormula('"a""b"')).toMatchObject({ type: "string", value: 'a"b' });
	});

	it("parses bare TRUE/FALSE as booleans but TRUE() as a call", () => {
		expect(parseFormula("TRUE")).toMatchObject({ type: "boolean", value: true });
		expect(parseFormula("FALSE")).toMatchObject({ type: "boolean", value: false });
		expect(parseFormula("TRUE()")).toMatchObject({ type: "call", name: "TRUE", args: [] });
	});

	it("parses all eight error literals as error values", () => {
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
			expect(parseFormula(e)).toEqual({ type: "error", value: e });
		}
	});
});

describe("parseFormula — references", () => {
	it("parses cells with $ markers preserved", () => {
		expect(parseFormula("$A$1")).toEqual({ type: "cell", ref: "$A$1" });
	});

	it("parses ranges, whole-column and whole-row (bare and absolute)", () => {
		expect(rp("A1:B2")).toBe("(A1:B2)");
		expect(rp("A:A")).toBe("(A:A)");
		expect(rp("1:1")).toBe("(1:1)");
		expect(rp("$A:$B")).toBe("($A:$B)");
		expect(rp("$2:$5")).toBe("($2:$5)");
	});

	it("attaches an unquoted sheet name to a reference", () => {
		expect(parseFormula("Sheet1!A1")).toEqual({
			type: "cell",
			ref: "A1",
			sheet: { name: "Sheet1" },
		});
	});

	it("attaches a quoted sheet name and unescapes it", () => {
		expect(parseFormula("'My Sheet'!A1")).toMatchObject({ sheet: { name: "My Sheet" } });
		expect(parseFormula("'O''Brien'!A1")).toMatchObject({ sheet: { name: "O'Brien" } });
	});

	it("parses 3-D sheet spans, quoted and unquoted", () => {
		expect(parseFormula("Sheet1:Sheet3!A1")).toMatchObject({
			sheet: { name: "Sheet1", toName: "Sheet3" },
		});
		expect(parseFormula("'Sheet 1:Sheet 3'!A1")).toMatchObject({
			sheet: { name: "Sheet 1", toName: "Sheet 3" },
		});
	});

	it("keeps the sheet on the left of a sheet-qualified range", () => {
		expect(rp("Sheet1!A1:B2")).toBe("(Sheet1!A1:B2)");
	});

	it("parses a reference union inside grouping parens", () => {
		expect(rp("(A1,A3,B5)")).toBe("(A1,A3,B5)");
	});

	it("distinguishes a grouping paren from a union", () => {
		expect(parseFormula("(A1+B1)")).toMatchObject({ type: "group" });
		expect(parseFormula("(A1,B1)")).toMatchObject({ type: "union" });
	});
});

describe("parseFormula — calls", () => {
	it("parses arguments and full expressions per argument", () => {
		expect(rp('IF(A1>0,"y","n")')).toBe('IF((A1>0),"y","n")');
	});

	it("keeps a top-level comma as an argument separator, not a union", () => {
		expect(parseFormula("SUM(A1,A2)")).toMatchObject({ type: "call", name: "SUM" });
		expect(parseFormula("SUM(A1,A2)").type).toBe("call");
	});

	it("represents an omitted argument as an empty node", () => {
		const ast = parseFormula("SUM(1,,2)");
		expect(ast).toMatchObject({ type: "call" });
		if (ast.type === "call") expect(ast.args[1]).toEqual({ type: "empty" });
	});

	it("nests a union used as a single argument", () => {
		expect(rp("SUM((A1,A2),3)")).toBe("SUM((A1,A2),3)");
	});

	it("accepts exactly 255 arguments but rejects 256", () => {
		const args255 = Array.from({ length: 255 }, () => "1").join(",");
		const args256 = Array.from({ length: 256 }, () => "1").join(",");
		expect(() => parseFormula(`SUM(${args255})`)).not.toThrow();
		expectCode(`SUM(${args256})`, "too-many-args");
	});
});

describe("parseFormula — array constants", () => {
	it("parses rows (;) and columns (,)", () => {
		expect(rp("{1,2;3,4}")).toBe("{1,2;3,4}");
	});

	it("allows signed literals, strings, booleans and errors", () => {
		expect(rp("{-1,2;TRUE,#REF!}")).toBe("{(-1),2;TRUE,#REF!}");
		expect(rp('{"a","b"}')).toBe('{"a","b"}');
	});

	it("rejects a reference or call inside an array constant", () => {
		expectCode("{A1}", "parse-error");
		expectCode("{SUM(1)}", "parse-error");
	});
});

describe("parseFormula — operators", () => {
	it("parses postfix percent and spill", () => {
		expect(rp("50%")).toBe("(50%)");
		expect(rp("A1#")).toBe("(A1#)");
	});

	it("parses implicit-intersection @ over a whole reference", () => {
		expect(rp("@A1:B2")).toBe("(@(A1:B2))");
	});

	it("tolerates a leading = (display form)", () => {
		expect(parseFormula("=A1")).toEqual({ type: "cell", ref: "A1" });
	});
});

describe("parseFormula — opaque nodes", () => {
	it("keeps a structured (table) reference as its verbatim source", () => {
		expect(parseFormula("Table1[@Amount]")).toEqual({
			type: "structured",
			source: "Table1[@Amount]",
		});
		expect(parseFormula("Table1[[#Data],[Amt]]")).toEqual({
			type: "structured",
			source: "Table1[[#Data],[Amt]]",
		});
	});

	it("keeps an external-workbook reference as its verbatim source", () => {
		expect(parseFormula("[1]Sheet1!A1")).toEqual({ type: "external", source: "[1]Sheet1!A1" });
		expect(parseFormula("'[Book.xlsx]Sheet1'!A1")).toEqual({
			type: "external",
			source: "'[Book.xlsx]Sheet1'!A1",
		});
	});

	it("treats a quoted drive-path or URL workbook as external, not a 3-D span", () => {
		// The ':' in `C:\` / `https:` must not be mistaken for the 3-D sheet-span separator.
		expect(parseFormula("'C:\\dir\\[Book.xlsx]Sheet1'!A1")).toMatchObject({ type: "external" });
		expect(parseFormula("'https://h/[Book.xlsx]Sheet1'!A1")).toEqual({
			type: "external",
			source: "'https://h/[Book.xlsx]Sheet1'!A1",
		});
	});
});

// Regressions pinned from the F8.1 adversarial review (parser lens). Each was a silent wrong tree
// or a wrong rejection before the fix.
describe("parseFormula — adversarial-review regressions (F8.1)", () => {
	it("binds a spill '#' to the reference it follows, not the whole range", () => {
		// SPILL_BP was below the ':' range op, so the right endpoint's '#' was torn off and wrapped
		// around the entire range. '#' is part of the reference — it must bind to each endpoint.
		expect(rp("A1#:B1#")).toBe("((A1#):(B1#))");
		expect(rp("A1:A5#")).toBe("(A1:(A5#))");
		expect(rp("A1#:A5")).toBe("((A1#):A5)");
		expect(rp("Sheet1!A1#:B1#")).toBe("((Sheet1!A1#):(B1#))");
		expect(rp("SUM(A1#:B1#)")).toBe("SUM(((A1#):(B1#)))");
		// '%' stays BELOW ':' (Excel: ':' > '%'), so a range percent wraps the whole range.
		expect(rp("A1:A5%")).toBe("((A1:A5)%)");
	});

	it("parses the deleted-sheet reference #REF!!A1 (sheet name became a #REF! error)", () => {
		expect(parseFormula("#REF!!A1")).toEqual({
			type: "cell",
			ref: "A1",
			sheet: { name: "#REF!" },
		});
		expect(rp("#REF!!A1:B2")).toBe("(#REF!!A1:B2)");
		expect(rp("SUM(#REF!!A1:A5)")).toBe("SUM((#REF!!A1:A5))");
		// A bare #REF! (not a deleted-sheet prefix) is still an error value.
		expect(parseFormula("#REF!")).toEqual({ type: "error", value: "#REF!" });
		// And the deleted-CELL form keeps working.
		expect(parseFormula("Sheet1!#REF!")).toEqual({ type: "error", value: "#REF!" });
	});
});

// Assert that parsing throws a FormulaError carrying the expected code.
function expectCode(formula: string, code: FormulaErrorCode): void {
	try {
		parseFormula(formula);
	} catch (e) {
		expect(e).toBeInstanceOf(FormulaError);
		if (e instanceof FormulaError) expect(e.code).toBe(code);
		return;
	}
	throw new Error(`expected ${JSON.stringify(formula)} to throw ${code}`);
}

describe("parseFormula — typed failures", () => {
	it.each([
		["", "parse-error"],
		["=", "parse-error"],
		["1+", "parse-error"],
		["(1", "parse-error"],
		["A1 A2", "parse-error"], // space-intersection is a named v0.8 exclusion
		["!A1", "parse-error"],
		["$", "parse-error"],
		['"abc', "parse-error"], // unterminated string
		["Table1[abc", "parse-error"], // unbalanced bracket
	])("%s → %s", (formula, code) => {
		expectCode(formula, code as FormulaErrorCode);
	});

	it("reports the character position of the failure", () => {
		try {
			parseFormula("1+*2");
			throw new Error("expected a throw");
		} catch (e) {
			expect(e).toBeInstanceOf(FormulaError);
			if (e instanceof FormulaError) expect(e.position).toBe(2);
		}
	});
});

describe("parseFormula — adversarial input is bounded and typed", () => {
	it("rejects a deep paren bomb with depth-exceeded, not a RangeError", () => {
		const bomb = `${"(".repeat(4096)}1${")".repeat(4096)}`;
		expectCode(bomb, "depth-exceeded");
	});

	it("rejects a deep unary chain with depth-exceeded", () => {
		expectCode(`${"-".repeat(5000)}1`, "depth-exceeded");
	});

	it("rejects over-deep function nesting with depth-exceeded (64-level limit)", () => {
		let f = "1";
		for (let i = 0; i < 100; i++) f = `SUM(${f})`;
		expectCode(f, "depth-exceeded");
	});

	it("parses a long flat chain iteratively (no stack growth)", () => {
		const chain = Array.from({ length: 20000 }, (_, i) => i).join("+");
		expect(() => parseFormula(chain)).not.toThrow();
	});

	it("parses a full-grid range without materializing it", () => {
		expect(rp("SUM(A1:XFD1048576)")).toBe("SUM((A1:XFD1048576))");
	});
});

// A cross-implementation corpus: every one of these parses in the Python `formulas` library
// (1.3.4) and must parse here too. It includes every distinct formula shape used by the F5.4
// shared-formula fixture/tests plus a spread of real-world functions and reference forms. The
// `formulas` agreement was verified out-of-tree (scratchpad venv); this pins that we accept them.
const REAL_WORLD = [
	"A1*2",
	"$A1+A$1",
	"SUM(A1:A10)",
	"SUM(B2:D2)",
	"A1+B1+C1",
	"Sheet2!A1+1",
	"'My Sheet'!A1*2",
	"LOG10(A1)+LN(B2)",
	"A1&Sheet1!B2",
	"IF(A1>0,B1,C1)",
	"AA10+AB11",
	"SUM(A1:A3)*Sheet2!B1",
	"A1<=B1",
	"-A1",
	"SUM(A:A)",
	"SUM(A:B)",
	"SUM($A:$B)",
	"SUM(A:$B)",
	"SUM($2:$5)",
	"VLOOKUP(A1,Data!A:C,3,0)",
	"A:A10",
	"COUNTIF(B:B,A1)",
	"A1+$B$2*Sheet1!C3",
	"ZZZ1+A1",
	'CONCATENATE(A9,",",B1)',
	'TEXT(B1,"0.00")',
	'IF(AND(A1>0,B1<10),"yes","no")',
	"INDEX(A1:C10,2,3)",
	"MATCH(A1,B1:B10,0)",
	'SUMIF(A1:A10,">5",B1:B10)',
	"IFERROR(A1/B1,0)",
	"ROUND(SUM(A1:A10)/COUNT(A1:A10),2)",
	"SUM({1,2;3,4})",
	'"line1"&CHAR(10)&"line2"',
	"MAX(0,MIN(100,A1))",
	"IFS(A1>0,1,A1<0,-1,TRUE,0)",
];

describe("parseFormula — real-world corpus (cross-checked vs Python `formulas`)", () => {
	it.each(REAL_WORLD)("parses %s", (formula) => {
		expect(() => parseFormula(formula)).not.toThrow();
	});
});
