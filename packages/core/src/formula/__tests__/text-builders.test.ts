import { describe, expect, it, onTestFinished, vi } from "vitest";
import { changeCase, concat, repeatText, replaceText, substitute, textjoin } from "../builtins";
import { applyBinary, joinText } from "../coerce";
import { ArrayView, capText, errorValue, MAX_TEXT } from "../value";

// F11.66: text that `&` or a built-in builds is #VALUE! past Excel's 32,767-character cell limit.

const VALUE = errorValue("#VALUE!");
const a = (n: number) => "a".repeat(n);
// A rope of 300 million characters: free to make, but joining two of them overflows a JS string.
const huge = a(300_000_000);

// The receivers of a String.prototype method's calls that were longer than the limit.
function longReceivers(method: "slice" | "toUpperCase" | "toLowerCase") {
	const spy = vi.spyOn(String.prototype, method);
	onTestFinished(() => spy.mockRestore());
	return () => spy.mock.contexts.filter((context) => String(context).length > MAX_TEXT);
}

describe("MAX_TEXT and capText", () => {
	it("is Excel's cell limit", () => {
		expect(MAX_TEXT).toBe(32767);
	});

	it("keeps text up to the limit and refuses longer text", () => {
		expect(capText("abc")).toBe("abc");
		expect(capText("")).toBe("");
		expect(capText(a(32767))).toBe(a(32767));
		expect(capText(a(32768))).toEqual(VALUE);
	});
});

describe("joinText", () => {
	it("joins both operands as text", () => {
		expect(joinText("a", 1)).toBe("a1");
		expect(joinText(true, null)).toBe("TRUE");
	});

	it("propagates the left error first, then the right", () => {
		expect(joinText(errorValue("#N/A"), errorValue("#DIV/0!"))).toEqual(errorValue("#N/A"));
		expect(joinText("a", errorValue("#DIV/0!"))).toEqual(errorValue("#DIV/0!"));
	});

	it("propagates an error before it looks at the length", () => {
		expect(joinText(a(40000), errorValue("#N/A"))).toEqual(errorValue("#N/A"));
		expect(joinText(errorValue("#N/A"), a(40000))).toEqual(errorValue("#N/A"));
	});

	it("allows a result of exactly 32,767 characters and refuses one more", () => {
		expect(joinText(a(32766), "b")).toBe(`${a(32766)}b`);
		expect(joinText(a(32767), "")).toBe(a(32767));
		expect(joinText(a(32767), "b")).toEqual(VALUE);
		expect(joinText("", a(32768))).toEqual(VALUE);
	});

	it("decides from the lengths, so text too long for JavaScript never throws", () => {
		expect(joinText(huge, huge)).toEqual(VALUE);
		expect(applyBinary("&", huge, huge)).toEqual(VALUE);
	});
});

describe("concat", () => {
	it("joins scalars, and a range's cells in row-major order", () => {
		expect(concat(["a", 1, true])).toBe("a1TRUE");
		expect(
			concat([
				new ArrayView([
					["a", "b"],
					["c", "d"],
				]),
				"e",
			]),
		).toBe("abcde");
	});

	it("propagates an error from a scalar or a range", () => {
		expect(concat(["a", errorValue("#N/A")])).toEqual(errorValue("#N/A"));
		expect(concat([new ArrayView([["a", errorValue("#DIV/0!")]])])).toEqual(
			errorValue("#DIV/0!"),
		);
	});

	it("allows exactly 32,767 characters and refuses one more, within a range too", () => {
		expect(concat([a(32766), "b"])).toBe(`${a(32766)}b`);
		expect(concat([new ArrayView([[a(32766), "b"]])])).toBe(`${a(32766)}b`);
		expect(concat([a(32767), "b"])).toEqual(VALUE);
		expect(concat([new ArrayView([[a(32767), "b"]])])).toEqual(VALUE);
	});

	it("still reads the rest of the argument that passes the limit, so an error in it wins", () => {
		expect(concat([new ArrayView([[a(32767), "b", errorValue("#N/A")]])])).toEqual(
			errorValue("#N/A"),
		);
		expect(concat([a(32767), "b", errorValue("#N/A")])).toEqual(VALUE);
	});

	it("never builds text too long for JavaScript", () => {
		expect(concat([huge, huge])).toEqual(VALUE);
		expect(concat([new ArrayView([[huge, huge]])])).toEqual(VALUE);
	});
});

describe("textjoin", () => {
	it("joins the items with the delimiter, skipping empty values only when asked", () => {
		expect(textjoin([",", true, "a", "", "b"])).toBe("a,b");
		expect(textjoin([",", false, "a", "", "b"])).toBe("a,,b");
		expect(textjoin(["-", true, new ArrayView([["a", null, "b"]]), "c"])).toBe("a-b-c");
		expect(textjoin([",", true])).toBe("");
	});

	it("propagates an error from the delimiter, ignore_empty or an item", () => {
		expect(textjoin([errorValue("#N/A"), true, "a"])).toEqual(errorValue("#N/A"));
		expect(textjoin([",", "yes", "a"])).toEqual(VALUE);
		expect(textjoin([",", true, "a", errorValue("#DIV/0!")])).toEqual(errorValue("#DIV/0!"));
	});

	it("counts every delimiter: exactly 32,767 characters is allowed, one more is not", () => {
		expect(textjoin(["--", true, a(32764), "b"])).toBe(`${a(32764)}--b`);
		expect(textjoin(["--", true, a(32765), "b"])).toEqual(VALUE);
		expect(textjoin([a(32767), true, "b"])).toBe("b");
		expect(textjoin([a(32766), false, "b", ""])).toBe(`b${a(32766)}`);
		expect(textjoin([a(32766), true, "b", "c"])).toEqual(VALUE);
		expect(textjoin([",", true, new ArrayView([[a(32766), "b"]])])).toEqual(VALUE);
	});

	it("still reads every item past the limit, so an error in any of them wins", () => {
		expect(textjoin([",", true, a(32767), "b", errorValue("#N/A")])).toEqual(
			errorValue("#N/A"),
		);
		expect(textjoin([",", true, new ArrayView([[a(32767), "b", errorValue("#N/A")]])])).toEqual(
			errorValue("#N/A"),
		);
	});

	it("never builds text too long for JavaScript", () => {
		expect(textjoin([",", true, huge, huge])).toEqual(VALUE);
		expect(textjoin([huge, false, "", "", ""])).toEqual(VALUE);
	});
});

describe("substitute", () => {
	it("replaces every occurrence, or only the given instance", () => {
		expect(substitute(["a-b-c", "-", "+"])).toBe("a+b+c");
		expect(substitute(["a-b-c", "-", "+", 2])).toBe("a-b+c");
		expect(substitute(["a-b-c", "-", "+", 3])).toBe("a-b-c");
		expect(substitute(["abc", "", "x"])).toBe("abc");
	});

	it("propagates an argument's error and refuses a bad instance", () => {
		expect(substitute([errorValue("#N/A"), "a", "b"])).toEqual(errorValue("#N/A"));
		expect(substitute(["abc", errorValue("#DIV/0!"), "b"])).toEqual(errorValue("#DIV/0!"));
		expect(substitute(["abc", "a", errorValue("#NUM!")])).toEqual(errorValue("#NUM!"));
		expect(substitute(["a-b-c", "-", "+", "two"])).toEqual(VALUE);
		expect(substitute(["a-b-c", "-", "+", 0])).toEqual(VALUE);
		expect(substitute([a(40000), errorValue("#N/A"), "b"])).toEqual(errorValue("#N/A"));
	});

	it("allows exactly 32,767 characters and refuses one more, at every instance", () => {
		expect(substitute([a(32767), "a", "b"])).toBe("b".repeat(32767));
		expect(substitute([`${a(32766)}x`, "x", "yy"])).toEqual(VALUE);
		expect(substitute([`${a(32765)}x`, "x", "yy"])).toBe(`${a(32765)}yy`);
		expect(substitute([`${a(32765)}x`, "x", "yy", 1])).toBe(`${a(32765)}yy`);
		expect(substitute([`${a(32766)}x`, "x", "yy", 1])).toEqual(VALUE);
	});

	it("counts the occurrences it replaces, which never overlap", () => {
		// 21,845 a's hold 10,922 pairs side by side, so the result is 21,845 + 10,922 = 32,767 long.
		expect(substitute([a(21845), "aa", "aaa"])).toBe(`${"aaa".repeat(10922)}a`);
		expect(substitute([a(21846), "aa", "aaa"])).toEqual(VALUE);
	});

	it("returns the text as it is when nothing is replaced, however long, as before", () => {
		expect(substitute([a(32768), "", "x"])).toBe(a(32768));
		expect(substitute([a(32768), "x", "y", 1])).toBe(a(32768));
		expect(substitute([a(32768), "x", "y"])).toBe(a(32768));
		expect(substitute(["abc", "x", "y"])).toBe("abc");
		expect(substitute([huge, "", "x"]) === huge).toBe(true);
	});

	it("caps text it builds even when the result reads the same", () => {
		expect(substitute([a(32768), "a", "a"])).toEqual(VALUE);
		expect(substitute([a(32768), "a", "a", 1])).toEqual(VALUE);
	});

	it("works out the length before building the text", () => {
		const long = longReceivers("slice");
		expect(substitute([a(40000), "a", "bb", 1])).toEqual(VALUE);
		expect(long()).toEqual([]);
		// The entry's case: building this first would overflow a JavaScript string.
		expect(substitute([a(32767), "a", "b".repeat(32767)])).toEqual(VALUE);
	});
});

describe("replaceText", () => {
	it("replaces num_chars characters from start_num on", () => {
		expect(replaceText(["abcdef", 2, 3, "X"])).toBe("aXef");
		expect(replaceText(["abc", 10, 2, "X"])).toBe("abcX");
		expect(replaceText(["abc", 2, 10, "X"])).toBe("aX");
	});

	it("propagates an argument's error and refuses a bad position", () => {
		expect(replaceText([errorValue("#N/A"), 1, 1, "x"])).toEqual(errorValue("#N/A"));
		expect(replaceText(["abc", "one", 1, "x"])).toEqual(VALUE);
		expect(replaceText(["abc", 1, "one", "x"])).toEqual(VALUE);
		expect(replaceText(["abc", 1, 1, errorValue("#DIV/0!")])).toEqual(errorValue("#DIV/0!"));
		expect(replaceText(["abc", 0, 1, "x"])).toEqual(VALUE);
		expect(replaceText(["abc", 1, -1, "x"])).toEqual(VALUE);
		expect(replaceText([a(40000), 1, 1, errorValue("#N/A")])).toEqual(errorValue("#N/A"));
	});

	it("allows exactly 32,767 characters and refuses one more, wherever start_num falls", () => {
		expect(replaceText([a(32767), 1, 1, "b"])).toBe(`b${a(32766)}`);
		expect(replaceText([a(32767), 1, 0, "b"])).toEqual(VALUE);
		expect(replaceText([a(32766), 40000, 1, "b"])).toBe(`${a(32766)}b`);
		expect(replaceText([a(32767), 40000, 0, "b"])).toEqual(VALUE);
		expect(replaceText([a(32767), 32767, 5, "bc"])).toEqual(VALUE);
		expect(replaceText([a(32767), 32766, 5, "bc"])).toBe(`${a(32765)}bc`);
	});

	it("works out the length before building the text", () => {
		const long = longReceivers("slice");
		expect(replaceText([a(40000), 1, 0, "b"])).toEqual(VALUE);
		expect(long()).toEqual([]);
	});
});

describe("changeCase", () => {
	it("upper- or lower-cases the text", () => {
		expect(changeCase(["Abc"], "upper")).toBe("ABC");
		expect(changeCase(["Abc"], "lower")).toBe("abc");
		expect(changeCase([errorValue("#N/A")], "upper")).toEqual(errorValue("#N/A"));
	});

	it("allows exactly 32,767 characters and refuses one more, after case mapping grows it", () => {
		expect(changeCase([a(32767)], "upper")).toBe("A".repeat(32767));
		expect(changeCase([`${"ß".repeat(16383)}a`], "upper")).toBe(`${"SS".repeat(16383)}A`);
		expect(changeCase(["ß".repeat(16384)], "upper")).toEqual(VALUE);
		// İ (U+0130) lower-cases to i and a combining dot (U+0307).
		expect(changeCase([`${"\u0130".repeat(16383)}a`], "lower")).toBe(
			`${"i\u0307".repeat(16383)}a`,
		);
		expect(changeCase(["\u0130".repeat(16384)], "lower")).toEqual(VALUE);
	});

	it("refuses longer input without case-mapping it", () => {
		const upper = longReceivers("toUpperCase");
		const lower = longReceivers("toLowerCase");
		expect(changeCase([a(40000)], "upper")).toEqual(VALUE);
		expect(changeCase([a(40000)], "lower")).toEqual(VALUE);
		expect([...upper(), ...lower()]).toEqual([]);
	});

	it("relies on case mapping never shortening text, which holds for every code point", () => {
		const shorter: number[] = [];
		for (let point = 0; point <= 0x10ffff; point++) {
			if (point >= 0xd800 && point <= 0xdfff) continue;
			const text = String.fromCodePoint(point);
			if (text.toUpperCase().length < text.length || text.toLowerCase().length < text.length)
				shorter.push(point);
		}
		expect(shorter).toEqual([]);
	});
});

describe("repeatText", () => {
	it("repeats the text", () => {
		expect(repeatText(["ab", 3])).toBe("ababab");
		expect(repeatText(["ab", 0])).toBe("");
	});

	it("repeats empty text to empty text for any count, an infinite one too", () => {
		expect(repeatText(["", 1e300])).toBe("");
		expect(repeatText(["", Number.POSITIVE_INFINITY])).toBe("");
		expect(repeatText([null, Number.POSITIVE_INFINITY])).toBe("");
		expect(repeatText(["a", Number.POSITIVE_INFINITY])).toEqual(VALUE);
		expect(repeatText(["", Number.NEGATIVE_INFINITY])).toEqual(VALUE);
	});

	it("propagates an argument's error and refuses a negative count", () => {
		expect(repeatText([errorValue("#N/A"), 2])).toEqual(errorValue("#N/A"));
		expect(repeatText(["ab", "two"])).toEqual(VALUE);
		expect(repeatText(["ab", -1])).toEqual(VALUE);
	});

	it("allows exactly 32,767 characters and refuses one more", () => {
		expect(repeatText(["a", 32767])).toBe(a(32767));
		expect(repeatText(["ab", 16384])).toEqual(VALUE);
		expect(repeatText(["a", 32768])).toEqual(VALUE);
	});
});
