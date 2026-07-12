import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { type EvaluateOptions, evaluateCell, evaluateWorkbook } from "../eval";
import { type EvalValue, errorValue } from "../value";

// The built-in function library (F8.3), driven end to end (writeXlsx → openXlsx → evaluateWorkbook).
// Expected values are DOCUMENTED Excel behavior, which is the authority. Where the Python `formulas`
// oracle diverges from Excel (SUM/COUNT/AVERAGE range coercion, TRIM internal-space collapse) we pin
// Excel; those divergences are asserted deliberately below and called out in comments.

const f = (formula: string): CellInput => ({ formula, value: 0 });
const err = (code: Parameters<typeof errorValue>[0]) => errorValue(code);

async function evalSheet(
	rows: readonly (readonly CellInput[])[],
	options?: EvaluateOptions,
): Promise<(ref: string) => EvalValue | undefined> {
	const bytes = await writeXlsx({ sheets: [{ name: "S", rows }] });
	const book = await openXlsx(bytes);
	const result = await evaluateWorkbook(book, options);
	return (ref) => result.get("S", ref);
}

describe("builtins — aggregates follow Excel range rules (text/logical ignored in a reference)", () => {
	it("SUM/COUNT/AVERAGE/MIN/MAX/COUNTA/COUNTBLANK over a mixed range", async () => {
		// A1..E1 = 10, "5"(text), TRUE, 20, (blank). Excel ignores the text and the boolean in a range.
		const g = await evalSheet([
			[10, "5", true, 20, null],
			[
				f("SUM(A1:E1)"),
				f("COUNT(A1:E1)"),
				f("AVERAGE(A1:E1)"),
				f("MAX(A1:E1)"),
				f("MIN(A1:E1)"),
			],
			[f("COUNTA(A1:E1)"), f("COUNTBLANK(A1:E1)")],
		]);
		expect(g("A2")).toBe(30); // 10 + 20 (text "5" and TRUE ignored — Excel, not the oracle's 35)
		expect(g("B2")).toBe(2); // only 10 and 20 are numbers
		expect(g("C2")).toBe(15); // (10 + 20) / 2
		expect(g("D2")).toBe(20);
		expect(g("E2")).toBe(10);
		expect(g("A3")).toBe(4); // COUNTA: 4 non-blank
		expect(g("B3")).toBe(1); // COUNTBLANK: E1 only
	});

	it("literal arguments coerce (SUM/COUNT), unlike cells inside a range", async () => {
		const g = await evalSheet([
			[f('SUM(10,"5",TRUE)'), f('COUNT(10,"5",TRUE,"x")'), f("AVERAGE(2,4)")],
		]);
		expect(g("A1")).toBe(16); // 10 + 5 + 1
		expect(g("B1")).toBe(3); // 10, "5", TRUE counted; "x" not
		expect(g("C1")).toBe(3);
	});

	it("AVERAGE: literal text → #VALUE!, all-text range → #DIV/0!; range errors propagate", async () => {
		// The literal-vs-range asymmetry, oracle-confirmed: coercing a literal text arg fails (#VALUE!),
		// while text inside a range is ignored — leaving no numbers → #DIV/0!.
		const g = await evalSheet([
			[f('AVERAGE("a","b")'), f("SUM(A2:B2)"), f("AVERAGE(A3:B3)")],
			[f("1/0"), 5],
			["x", "y"],
		]);
		expect(g("A1")).toEqual(err("#VALUE!"));
		expect(g("B1")).toEqual(err("#DIV/0!")); // A2 is #DIV/0!, propagates through SUM
		expect(g("C1")).toEqual(err("#DIV/0!")); // A3:B3 all text → no numbers
	});
});

describe("builtins — math", () => {
	it("rounding family, INT, ABS, MOD, SIGN, TRUNC", async () => {
		const g = await evalSheet([
			[
				f("ROUND(2.5,0)"),
				f("ROUNDUP(2.1,0)"),
				f("ROUNDDOWN(2.9,0)"),
				f("INT(-2.5)"),
				f("ABS(-7)"),
				f("MOD(-3,2)"),
				f("SIGN(-5)"),
				f("TRUNC(3.99)"),
				f("ROUND(-2.5,0)"),
				f("ROUND(3.14159,2)"),
			],
		]);
		expect(g("A1")).toBe(3);
		expect(g("B1")).toBe(3);
		expect(g("C1")).toBe(2);
		expect(g("D1")).toBe(-3);
		expect(g("E1")).toBe(7);
		expect(g("F1")).toBe(1); // sign of divisor
		expect(g("G1")).toBe(-1);
		expect(g("H1")).toBe(3);
		expect(g("I1")).toBe(-3); // half away from zero
		expect(g("J1")).toBe(3.14);
	});

	it("POWER/SQRT/EXP/LN/LOG/LOG10/PI and their domain errors", async () => {
		const g = await evalSheet([
			[
				f("POWER(2,10)"),
				f("SQRT(16)"),
				f("SQRT(-1)"),
				f("LN(1)"),
				f("LOG(8,2)"),
				f("LOG10(1000)"),
				f("ROUND(PI(),5)"),
				f("MOD(10,0)"),
			],
		]);
		expect(g("A1")).toBe(1024);
		expect(g("B1")).toBe(4);
		expect(g("C1")).toEqual(err("#NUM!"));
		expect(g("D1")).toBe(0);
		expect(g("E1")).toBe(3);
		expect(g("F1")).toBe(3);
		expect(g("G1")).toBe(3.14159);
		expect(g("H1")).toEqual(err("#DIV/0!"));
	});

	it("MEDIAN/LARGE/SMALL/SUMPRODUCT", async () => {
		const g = await evalSheet([
			[3, 1, 2, 10],
			[
				f("MEDIAN(A1:D1)"),
				f("LARGE(A1:D1,1)"),
				f("SMALL(A1:D1,2)"),
				f("SUMPRODUCT(A1:D1,A1:D1)"),
			],
			[f("LARGE(A1:D1,9)")],
		]);
		expect(g("A2")).toBe(2.5); // median of 1,2,3,10
		expect(g("B2")).toBe(10);
		expect(g("C2")).toBe(2);
		expect(g("D2")).toBe(114); // 9+1+4+100
		expect(g("A3")).toEqual(err("#NUM!"));
	});
});

describe("builtins — logical", () => {
	it("IF is lazy; the untaken branch (and its errors) never evaluate", async () => {
		const g = await evalSheet([
			[1, 0],
			[f("IF(A1>0,A1,B1/0)"), f("IF(B1>0,A1/0,99)"), f('IF("TRUE",1,2)')],
		]);
		expect(g("A2")).toBe(1); // no #DIV/0! from untaken branch
		expect(g("B2")).toBe(99);
		expect(g("C2")).toEqual(err("#VALUE!")); // text condition is not coerced
	});

	it("IFERROR/IFNA/AND/OR/NOT/XOR/IFS/SWITCH/CHOOSE", async () => {
		const g = await evalSheet([
			[1, 2, 3],
			[
				f("IFERROR(1/0,99)"),
				f("IFNA(NA(),7)"),
				f("IFNA(1/0,7)"),
				f("AND(A1:C1)"),
				f("OR(FALSE,0)"),
			],
			[
				f("NOT(FALSE)"),
				f("XOR(TRUE,TRUE,TRUE)"),
				f("IFS(FALSE,1,TRUE,2)"),
				f('SWITCH(2,1,"a",2,"b","z")'),
			],
			[f('CHOOSE(3,"x","y","z")'), f("CHOOSE(9,1,2)")],
		]);
		expect(g("A2")).toBe(99);
		expect(g("B2")).toBe(7);
		expect(g("C2")).toEqual(err("#DIV/0!")); // IFNA only catches #N/A
		expect(g("D2")).toBe(true); // all nonzero
		expect(g("E2")).toBe(false);
		expect(g("A3")).toBe(true);
		expect(g("B3")).toBe(true); // odd count of TRUE
		expect(g("C3")).toBe(2);
		expect(g("D3")).toBe("b");
		expect(g("A4")).toBe("z");
		expect(g("B4")).toEqual(err("#VALUE!")); // index out of range
	});
});

describe("builtins — lookup & reference", () => {
	it("VLOOKUP exact & approximate, HLOOKUP, MATCH, INDEX", async () => {
		const g = await evalSheet([
			[1, "one"],
			[2, "two"],
			[3, "three"],
			[10, "ten"],
			[
				f("VLOOKUP(3,A1:B4,2,FALSE)"),
				f("VLOOKUP(4,A1:B4,2,TRUE)"),
				f("VLOOKUP(99,A1:B4,2,FALSE)"),
				f("MATCH(3,A1:A4,0)"),
				f("MATCH(4,A1:A4,1)"),
				f("INDEX(A1:B4,2,2)"),
				f("INDEX(A1:A4,9)"),
				f("VLOOKUP(3,A1:B4,5,FALSE)"),
			],
		]);
		expect(g("A5")).toBe("three"); // exact
		expect(g("B5")).toBe("three"); // approx: largest ≤ 4 is 3
		expect(g("C5")).toEqual(err("#N/A")); // exact miss
		expect(g("D5")).toBe(3);
		expect(g("E5")).toBe(3); // approx match position
		expect(g("F5")).toBe("two");
		expect(g("G5")).toEqual(err("#REF!")); // out of range
		expect(g("H5")).toEqual(err("#REF!")); // col index beyond table width
	});

	it("MATCH type -1 needs descending data (early break); ascending data → #N/A", async () => {
		const g = await evalSheet([
			[30, 20, 10],
			[f("MATCH(25,A1:C1,-1)"), f("MATCH(5,A1:C1,-1)")],
			[10, 20, 30],
			[f("MATCH(25,A3:C3,-1)")], // ascending data with a descending search → #N/A
		]);
		expect(g("A2")).toBe(1); // smallest ≥ 25 in descending {30,20,10} is 30 → position 1
		expect(g("B2")).toBe(3); // smallest ≥ 5 is 10 → position 3
		expect(g("A4")).toEqual(err("#N/A"));
	});

	it("HLOOKUP, ROWS/COLUMNS, INDEX with blank → 0", async () => {
		const g = await evalSheet([
			[1, 2, 3],
			["a", "b", "c"],
			[f("HLOOKUP(2,A1:C2,2,FALSE)"), f("ROWS(A1:C2)"), f("COLUMNS(A1:C2)"), f("ROWS(A1)")],
		]);
		expect(g("A3")).toBe("b");
		expect(g("B3")).toBe(2);
		expect(g("C3")).toBe(3);
		expect(g("D3")).toBe(1); // single cell
	});
});

describe("builtins — conditional aggregates", () => {
	it("SUMIF/COUNTIF/AVERAGEIF/SUMIFS/COUNTIFS/AVERAGEIFS", async () => {
		const g = await evalSheet([
			[1, 5],
			[2, 6],
			[3, 7],
			[10, 8],
			[
				f('SUMIF(A1:A4,">2")'),
				f('SUMIF(A1:A4,">2",B1:B4)'),
				f('COUNTIF(A1:A4,">=3")'),
				f('AVERAGEIF(A1:A4,">1")'),
				f('SUMIFS(B1:B4,A1:A4,">1",A1:A4,"<10")'),
				f('COUNTIFS(A1:A4,">1",B1:B4,"<8")'),
				f('AVERAGEIFS(B1:B4,A1:A4,">=3")'),
			],
		]);
		expect(g("A5")).toBe(13); // 3 + 10
		expect(g("B5")).toBe(15); // B3 + B4 = 7 + 8
		expect(g("C5")).toBe(2); // 3, 10
		expect(g("D5")).toBe(5); // (2+3+10)/3
		expect(g("E5")).toBe(13); // B2+B3 (a=2,3; a<10 excludes 10) = 6+7
		expect(g("F5")).toBe(2); // a in {2,3} with b in {6,7} < 8
		expect(g("G5")).toBe(7.5); // (7+8)/2
	});

	it("COUNTIF with text and wildcards", async () => {
		const g = await evalSheet([
			["apple", "banana", "apricot", "cherry"],
			[f('COUNTIF(A1:D1,"a*")'), f('COUNTIF(A1:D1,"apple")'), f('COUNTIF(A1:D1,"<>apple")')],
		]);
		expect(g("A2")).toBe(2); // apple, apricot
		expect(g("B2")).toBe(1);
		expect(g("C2")).toBe(3); // banana, apricot, cherry
	});
});

describe("builtins — text", () => {
	it("CONCAT/LEN/LEFT/RIGHT/MID/TRIM/UPPER/LOWER/PROPER", async () => {
		const g = await evalSheet([
			[
				f('CONCAT("a",1,TRUE)'),
				f("LEN(123.5)"),
				f('LEFT("hello")'),
				f('RIGHT("hello",2)'),
				f('MID("hello",2,3)'),
				f('TRIM("  a   b  ")'),
				f('UPPER("aB")'),
				f('LOWER("aB")'),
				f('PROPER("hELLo wORLD")'),
			],
		]);
		expect(g("A1")).toBe("a1TRUE");
		expect(g("B1")).toBe(5); // "123.5"
		expect(g("C1")).toBe("h"); // default 1
		expect(g("D1")).toBe("lo");
		expect(g("E1")).toBe("ell");
		expect(g("F1")).toBe("a b"); // Excel collapses internal runs (oracle would keep "a   b")
		expect(g("G1")).toBe("AB");
		expect(g("H1")).toBe("ab");
		expect(g("I1")).toBe("Hello World");
	});

	it("SUBSTITUTE/REPLACE/FIND/SEARCH/VALUE/REPT/EXACT/CHAR/CODE/TEXTJOIN", async () => {
		const g = await evalSheet([
			[1, 2, 3],
			[
				f('SUBSTITUTE("aXbXc","X","-")'),
				f('SUBSTITUTE("aXbXc","X","-",2)'),
				f('REPLACE("abcdef",2,3,"XY")'),
				f('FIND("c","abcabc")'),
				f('SEARCH("B","aAbB")'),
				f('VALUE("12.5")'),
				f('REPT("ab",3)'),
				f('EXACT("a","A")'),
				f("CHAR(65)"),
				f('CODE("A")'),
				f('TEXTJOIN("-",TRUE,A1:C1)'),
			],
		]);
		expect(g("A2")).toBe("a-b-c");
		expect(g("B2")).toBe("aXb-c"); // only 2nd occurrence
		expect(g("C2")).toBe("aXYef");
		expect(g("D2")).toBe(3);
		expect(g("E2")).toBe(3); // case-insensitive: first b/B
		expect(g("F2")).toBe(12.5);
		expect(g("G2")).toBe("ababab");
		expect(g("H2")).toBe(false); // case-sensitive
		expect(g("I2")).toBe("A");
		expect(g("J2")).toBe(65);
		expect(g("K2")).toBe("1-2-3");
	});
});

describe("builtins — information", () => {
	it("IS-family, N, T, NA, ERROR.TYPE", async () => {
		const g = await evalSheet([
			[42, "hi", true, null],
			[
				f("ISNUMBER(A1)"),
				f("ISTEXT(B1)"),
				f("ISLOGICAL(C1)"),
				f("ISBLANK(D1)"),
				f("ISERROR(1/0)"),
				f("ISNA(NA())"),
				f("ISERR(NA())"),
				f("N(C1)"),
				f("T(B1)"),
				f("ERROR.TYPE(1/0)"),
			],
		]);
		expect(g("A2")).toBe(true);
		expect(g("B2")).toBe(true);
		expect(g("C2")).toBe(true);
		expect(g("D2")).toBe(true);
		expect(g("E2")).toBe(true);
		expect(g("F2")).toBe(true);
		expect(g("G2")).toBe(false); // ISERR excludes #N/A
		expect(g("H2")).toBe(1); // N(TRUE)
		expect(g("I2")).toBe("hi");
		expect(g("J2")).toBe(2); // #DIV/0! → 2
	});
});

describe("builtins — date & time", () => {
	it("DATE/YEAR/MONTH/DAY/WEEKDAY/DAYS/EDATE/EOMONTH/TIME", async () => {
		const g = await evalSheet([
			[
				f("DATE(2020,1,15)"),
				f("YEAR(DATE(2020,1,15))"),
				f("MONTH(DATE(2020,1,15))"),
				f("DAY(DATE(2020,1,15))"),
				f("WEEKDAY(DATE(2020,1,15))"),
				f("DAYS(DATE(2020,1,20),DATE(2020,1,15))"),
				f("DAY(EDATE(DATE(2020,1,31),1))"),
				f("DAY(EOMONTH(DATE(2020,1,15),0))"),
				f("HOUR(TIME(13,30,0))"),
			],
		]);
		expect(g("A1")).toBe(43845); // 2020-01-15 serial
		expect(g("B1")).toBe(2020);
		expect(g("C1")).toBe(1);
		expect(g("D1")).toBe(15);
		expect(g("E1")).toBe(4); // Wed = 4 (Sun=1)
		expect(g("F1")).toBe(5);
		expect(g("G1")).toBe(29); // Jan 31 + 1 month → Feb 29 (2020 leap), clamped
		expect(g("H1")).toBe(31); // end of January
		expect(g("I1")).toBe(13);
	});
});

describe("builtins — volatile gate & unknown function", () => {
	it("TODAY/NOW/RAND require injection and are deterministic with it", async () => {
		const bytes = await writeXlsx({
			sheets: [
				{
					name: "S",
					rows: [[f("TODAY()"), f("NOW()"), f("RAND()"), f("RANDBETWEEN(1,1)")]],
				},
			],
		});
		const book = await openXlsx(bytes);
		await expect(evaluateCell(book, "S", "A1")).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
		const now = () => new Date(Date.UTC(2020, 0, 15, 12, 0, 0));
		const r = await evaluateWorkbook(book, { now, random: () => 0.5 });
		expect(r.get("S", "A1")).toBe(43845); // TODAY floors the serial
		expect(r.get("S", "B1")).toBe(43845.5); // NOW keeps the time
		expect(r.get("S", "C1")).toBe(0.5);
		expect(r.get("S", "D1")).toBe(1);
	});

	it("unknown function → #NAME?, wrong arity → #VALUE!", async () => {
		const g = await evalSheet([[f("NOSUCHFN(1)"), f("SUM()"), f("ABS(1,2)")]]);
		expect(g("A1")).toEqual(err("#NAME?"));
		expect(g("B1")).toEqual(err("#VALUE!")); // SUM needs ≥1 arg
		expect(g("C1")).toEqual(err("#VALUE!")); // ABS takes exactly 1
	});
});

// Regressions pinned from the F8.3 adversarial review (11 distinct bugs, each oracle-confirmed).
describe("builtins — adversarial-review regressions (F8.3)", () => {
	it("COUNT ignores error values instead of propagating them", async () => {
		const g = await evalSheet([[f("COUNT(1/0,5)"), f("COUNT(NA(),1,2)"), f("SUM(1/0,5)")]]);
		expect(g("A1")).toBe(1); // the error is ignored; 5 counts
		expect(g("B1")).toBe(2); // #N/A ignored; 1 and 2 count
		expect(g("C1")).toEqual(err("#DIV/0!")); // SUM still propagates
	});

	it("CEILING/FLOOR handle a negative number with a positive significance", async () => {
		const g = await evalSheet([
			[f("CEILING(-2.5,1)"), f("FLOOR(-2.5,1)"), f("CEILING(2.5,-1)"), f("CEILING(-2.5,2)")],
		]);
		expect(g("A1")).toBe(-2);
		expect(g("B1")).toBe(-3);
		expect(g("C1")).toEqual(err("#NUM!")); // (positive number, negative significance) is the only reject
		expect(g("D1")).toBe(-2); // Microsoft docs example
	});

	it("SUMPRODUCT propagates an error even where the first factor is 0/blank", async () => {
		const g = await evalSheet([
			[0, f("1/0")], // A1=0, B1=#DIV/0!
			[1, 2], // A2=1, B2=2
			[f("SUMPRODUCT(A1:A2,B1:B2)")],
		]);
		expect(g("A3")).toEqual(err("#DIV/0!"));
	});

	it("WEEKDAY supports return types 11-17", async () => {
		const g = await evalSheet([
			[f("WEEKDAY(DATE(2020,1,15),11)"), f("WEEKDAY(DATE(2020,1,15),16)")],
		]);
		expect(g("A1")).toBe(3); // 2020-01-15 is Wed; type 11 (Mon=1) → 3
		expect(g("B1")).toBe(5); // type 16 (Sat=1) → 5
	});

	it("EXP overflow and POWER(0,negative) return typed errors, not Infinity", async () => {
		const g = await evalSheet([
			[f("EXP(1000)"), f("POWER(0,-1)"), f("TRUNC(5,400)"), f("ROUND(5,-400)")],
		]);
		expect(g("A1")).toEqual(err("#NUM!")); // was Infinity
		expect(g("B1")).toEqual(err("#DIV/0!")); // consistent with 0^-1
		expect(g("C1")).toBe(5); // huge digits → unchanged, was NaN
		expect(g("D1")).toBe(0); // rounding past the number → 0
	});

	it("*IF family counts blank cells that satisfy the criterion", async () => {
		const g = await evalSheet([
			["x", 10], // A1, B1
			[null, 20], // A2 blank, B2
			["y", 30], // A3, B3
			[null, 40], // A4 blank, B4
			[
				f('COUNTIF(A1:A4,"<>x")'),
				f('COUNTIF(A1:A4,"")'),
				f('SUMIF(A1:A4,"<>x",B1:B4)'),
				f('COUNTIFS(A1:A4,"<>x",B1:B4,">0")'),
				f('SUMIFS(B1:B4,A1:A4,"<>x")'),
			],
		]);
		expect(g("A5")).toBe(3); // A2, A3, A4 (both blanks match "<>x")
		expect(g("B5")).toBe(2); // the two blank cells
		expect(g("C5")).toBe(90); // B2+B3+B4 (blank-A rows included)
		expect(g("D5")).toBe(3); // COUNTIFS with a blank-excluding companion criterion
		expect(g("E5")).toBe(90); // single-criterion SUMIFS delegates to the blank-aware path
	});

	it("blank-aware SUMIF reshapes a larger value range to the criteria range (no over-count)", async () => {
		// Second-pass-review regression: the blank pass must clip the value range to the criteria
		// rectangle. Excel reshapes B1:B5 → B1:B3 (anchored top-left), so B4/B5 never participate.
		const g = await evalSheet([
			[1, 10],
			[1, 20],
			[1, 30],
			[null, 40],
			[null, 50],
			[
				f('SUMIF(A1:A3,"<>x",B1:B5)'),
				f('AVERAGEIF(A1:A3,"<>x",B1:B5)'),
				f('SUMIFS(B1:B5,A1:A3,"<>x")'),
			],
		]);
		expect(g("A6")).toBe(60); // B1+B2+B3, NOT B1:B5 (=150)
		expect(g("B6")).toBe(20); // (10+20+30)/3, not inflated
		expect(g("C6")).toBe(60); // single-pair SUMIFS delegates to the same reshaped path
	});

	it("VALUE accepts percent, thousands separators and a currency sign", async () => {
		const g = await evalSheet([[f('VALUE("50%")'), f('VALUE("1,000")'), f('VALUE("$5")')]]);
		expect(g("A1")).toBe(0.5);
		expect(g("B1")).toBe(1000);
		expect(g("C1")).toBe(5);
	});

	it("date functions reject out-of-range serials with #NUM!", async () => {
		const g = await evalSheet([
			[f("YEAR(100000000)"), f("EOMONTH(DATE(1900,1,15),-2)"), f("EDATE(1,1000000000)")],
		]);
		expect(g("A1")).toEqual(err("#NUM!")); // serial beyond 9999-12-31
		expect(g("B1")).toEqual(err("#NUM!")); // result before 1900
		expect(g("C1")).toEqual(err("#NUM!")); // absurd month count → non-finite
	});

	it("IFERROR/IFNA pass a multi-cell range through (not mistaken for an error)", async () => {
		// Milestone-review regression: scalarize reduces a range to a #VALUE! sentinel, which IFERROR
		// must NOT read as a genuine error — an aggregator around it still needs the whole range.
		const g = await evalSheet([
			[1, 2, 3],
			[
				f("SUM(IFERROR(A1:C1,0))"),
				f("MAX(IFERROR(A1:C1,99))"),
				f("IFERROR(A1:C1,-1)"),
				f("SUM(IFNA(A1:C1,0))"),
			],
		]);
		expect(g("A2")).toBe(6); // was 0 (fallback), now the summed range
		expect(g("B2")).toBe(3); // was 99
		expect(g("C2")).toBe(1); // scalar context → top-left of the range
		expect(g("D2")).toBe(6);
	});

	it("accepts a time-of-day on the last valid day (9999-12-31), rejecting only the next day", async () => {
		// Second-pass-review regression: the serial cap must compare the truncated DAY, so a datetime
		// like 2958465.5 (9999-12-31 noon) is valid while 2958466 (year 10000) is out of range.
		const g = await evalSheet([
			[f("YEAR(2958465.5)"), f("DAY(2958465.5)"), f("YEAR(2958466)")],
		]);
		expect(g("A1")).toBe(9999);
		expect(g("B1")).toBe(31);
		expect(g("C1")).toEqual(err("#NUM!"));
	});

	it("wildcard matching cannot catastrophically backtrack (ReDoS)", async () => {
		// The classic near-miss: `*a*a*…*a*b` against a long run of 'a's that never contains 'b'. A regex
		// translation backtracks exponentially and hangs; the linear glob matcher fails fast.
		const pattern = `*${"a*".repeat(24)}b`; // requires a 'b' after 24 a-runs
		const text = "a".repeat(50); // no 'b' — the match must fail
		const start = performance.now();
		const g = await evalSheet([[text], [f(`COUNTIF(A1,"${pattern}")`)]]);
		expect(g("A2")).toBe(0); // no match, and — crucially — it does not hang
		expect(performance.now() - start).toBeLessThan(2000);
	});
});
