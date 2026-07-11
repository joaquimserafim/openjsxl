import type { ArrayElement, BinaryOp, FormulaAst, SheetSpec, UnaryOp } from "./ast";
import { FormulaError } from "./errors";
import { type Token, tokenize } from "./lexer";

// A precedence-climbing parser over the token stream, producing the typed AST in `ast.ts`. It is
// recursive-descent, but every recursion path passes through `parseExpr`, whose depth guard bounds
// the JS call stack (M8 decision 7): a paren bomb, a unary chain, or deep call nesting all trip a
// typed `FormulaError` long before V8's own stack limit. Ranges are never materialized and argument
// counts are capped, so a hostile stored form costs bounded time and memory.

// Nesting/recursion guard. Excel limits FUNCTION nesting to 64 but allows far deeper parens; we cap
// the parser's own recursion so a `((((…))))` bomb inside the 8192-char stored form is a typed error,
// not a `RangeError`. 256 comfortably clears any real formula and stays well under the engine limit.
const MAX_DEPTH = 256;
const MAX_FUNCTION_DEPTH = 64;
const MAX_ARGS = 255;

// Prefix binding power for unary `-`/`+`/`@`. Above `^`(70) so `-2^2` is `(-2)^2 = 4` and above
// `%`(80) so `-2%` is `(-2)%`, both matching Excel's documented precedence (M8 decision 6). The
// range operator `:`(100) still binds tighter, so `-A1:B2` negates the whole range.
const PREFIX_BP = 90;
const PERCENT_BP = 80; // postfix `%` — below the `:` range op, so `A1:A5%` is `(A1:A5)%` (Excel: `:` > `%`)
// Postfix `#` (spilled range) — ABOVE the `:` range op (100), because `#` is part of the reference it
// follows: `A1:A5#` is `A1:(A5#)` and `A1#:B1#` binds `#` to BOTH endpoints, not the whole range.
const SPILL_BP = 110;

// Infix binding power, higher = tighter. `,` (union) and ` ` (intersection) are NOT here: union is
// handled structurally inside grouping parens, and intersection-by-space is a named v0.8 exclusion.
// Everything is left-associative (right min-bp = bp + 1).
function infixBp(t: Token): number {
	if (t.type !== "op") return -1;
	switch (t.value) {
		case ":":
			return 100;
		case "^":
			return 70;
		case "*":
		case "/":
			return 60;
		case "+":
		case "-":
			return 50;
		case "&":
			return 40;
		case "=":
		case "<>":
		case "<":
		case ">":
		case "<=":
		case ">=":
			return 30;
		default:
			return -1;
	}
}

function isBooleanLiteral(name: string): boolean {
	return name.length === 4 || name.length === 5 ? /^(?:true|false)$/i.test(name) : false;
}

class Parser {
	private pos = 0;
	private depth = 0;
	private funcDepth = 0;
	private readonly last: number;

	constructor(
		private readonly source: string,
		private readonly tokens: Token[],
	) {
		this.last = tokens.length - 1; // index of the trailing eof token
	}

	private peek(offset = 0): Token {
		const i = this.pos + offset;
		return this.tokens[i > this.last ? this.last : i] as Token;
	}

	private next(): Token {
		const t = this.tokens[this.pos] as Token;
		if (t.type !== "eof") this.pos++;
		return t;
	}

	private isOp(t: Token, value: string): boolean {
		return t.type === "op" && t.value === value;
	}

	private fail(code: "parse-error" | "unsupported", message: string, at: Token): never {
		throw new FormulaError(code, message, { position: at.start });
	}

	private expectOp(value: string): void {
		const t = this.peek();
		if (this.isOp(t, value)) {
			this.next();
			return;
		}
		this.fail("parse-error", `expected '${value}'`, t);
	}

	private enter(): void {
		if (++this.depth > MAX_DEPTH) {
			throw new FormulaError("depth-exceeded", "formula nesting too deep", {
				position: this.peek().start,
			});
		}
	}

	private leave(): void {
		this.depth--;
	}

	/** Parse the whole token stream, requiring it to be fully consumed. */
	parse(): FormulaAst {
		const ast = this.parseExpr(0);
		const t = this.peek();
		if (t.type !== "eof") this.fail("parse-error", `unexpected '${t.value}'`, t);
		return ast;
	}

	private parseExpr(minBp: number): FormulaAst {
		this.enter();
		let left = this.parseUnary();
		for (;;) {
			const t = this.peek();
			// Postfix operators bind to the operand already built.
			if (this.isOp(t, "%") && PERCENT_BP >= minBp) {
				this.next();
				left = { type: "postfix", op: "%", operand: left };
				continue;
			}
			if (this.isOp(t, "#") && SPILL_BP >= minBp) {
				this.next();
				left = { type: "postfix", op: "#", operand: left };
				continue;
			}
			const bp = infixBp(t);
			if (bp < minBp) break;
			this.next();
			const right = this.parseExpr(bp + 1); // left-assoc
			if (t.value === ":") {
				left = { type: "range", left, right };
			} else {
				left = { type: "binary", op: t.value as BinaryOp["op"], left, right };
			}
		}
		this.leave();
		return left;
	}

	private parseUnary(): FormulaAst {
		const t = this.peek();
		if (this.isOp(t, "-") || this.isOp(t, "+") || this.isOp(t, "@")) {
			this.next();
			const operand = this.parseExpr(PREFIX_BP);
			return { type: "unary", op: t.value as UnaryOp["op"], operand };
		}
		return this.parsePrimary();
	}

	private parsePrimary(): FormulaAst {
		const t = this.peek();
		switch (t.type) {
			case "num":
				this.next();
				return { type: "number", value: Number(t.value), raw: t.value };
			case "str":
				this.next();
				return { type: "string", value: t.value };
			case "err":
				// `#REF!!A1` — the stored form Excel writes when a formula's referenced SHEET is
				// deleted: the sheet name becomes the `#REF!` error, still followed by the `!`
				// separator. Treat the error as the sheet name so the reference parses (and degrades
				// at eval) rather than being rejected wholesale. Mirrors the deleted-CELL form
				// `Sheet1!#REF!`, which parseSheetCore already accepts after the `!`.
				if (this.isOp(this.peek(1), "!")) {
					this.next(); // err
					this.next(); // '!'
					return this.parseSheetCore({ name: t.value }, t.start, false);
				}
				this.next();
				return { type: "error", value: t.value };
			case "cell":
				this.next();
				return { type: "cell", ref: t.value };
			case "sheet":
				return this.parseQuotedSheetRef();
			case "bracket":
				return this.parseExternalBracketRef();
			case "name":
				return this.parseNameLike();
			case "op":
				if (t.value === "(") return this.parseGroupOrUnion();
				if (t.value === "{") return this.parseArray();
				return this.fail("parse-error", `unexpected '${t.value}'`, t);
			default:
				return this.fail("parse-error", "unexpected end of formula", t);
		}
	}

	// `( … )` — a grouping, OR a reference union when top-level commas appear: `(A1,A3,B5)`. The `,`
	// only means union here; inside a function call it is an argument separator (handled in parseCall).
	private parseGroupOrUnion(): FormulaAst {
		this.next(); // '('
		const first = this.parseExpr(0);
		if (this.isOp(this.peek(), ",")) {
			const items: FormulaAst[] = [first];
			while (this.isOp(this.peek(), ",")) {
				this.next();
				items.push(this.parseExpr(0));
			}
			this.expectOp(")");
			return { type: "union", items };
		}
		this.expectOp(")");
		return { type: "group", expr: first };
	}

	// `{1,2;3,4}` — an array constant. `,` separates columns, `;` separates rows; elements are
	// literals only (optionally signed), never references or calls.
	private parseArray(): FormulaAst {
		this.next(); // '{'
		const rows: ArrayElement[][] = [];
		let row: ArrayElement[] = [];
		for (;;) {
			row.push(this.parseArrayElement());
			const t = this.peek();
			if (this.isOp(t, ",")) {
				this.next();
				continue;
			}
			if (this.isOp(t, ";")) {
				this.next();
				rows.push(row);
				row = [];
				continue;
			}
			if (this.isOp(t, "}")) {
				this.next();
				rows.push(row);
				break;
			}
			this.fail("parse-error", "expected ',', ';' or '}' in array constant", t);
		}
		return { type: "array", rows };
	}

	private parseArrayElement(): ArrayElement {
		const t = this.peek();
		if (this.isOp(t, "-") || this.isOp(t, "+")) {
			this.next();
			return { type: "unary", op: t.value as "-" | "+", operand: this.parseArrayScalar() };
		}
		return this.parseArrayScalar();
	}

	private parseArrayScalar(): ArrayElement {
		const t = this.peek();
		switch (t.type) {
			case "num":
				this.next();
				return { type: "number", value: Number(t.value), raw: t.value };
			case "str":
				this.next();
				return { type: "string", value: t.value };
			case "err":
				this.next();
				return { type: "error", value: t.value };
			case "name":
				if (isBooleanLiteral(t.value)) {
					this.next();
					return { type: "boolean", value: /^true$/i.test(t.value) };
				}
				return this.fail("parse-error", "array constants may contain literals only", t);
			default:
				return this.fail("parse-error", "array constants may contain literals only", t);
		}
	}

	// A `name` token: a function call, a structured reference, a sheet-qualified reference (incl. a
	// 3-D span), a boolean literal, or a plain defined name — disambiguated by the following tokens.
	private parseNameLike(): FormulaAst {
		const nameTok = this.peek();
		const p2 = this.peek(1);

		if (this.isOp(p2, "(")) return this.parseCall(nameTok);
		if (p2.type === "bracket") {
			this.next();
			this.next();
			return { type: "structured", source: this.source.slice(nameTok.start, p2.end) };
		}
		if (this.isOp(p2, "!")) {
			this.next(); // name
			this.next(); // '!'
			return this.parseSheetCore({ name: nameTok.value }, nameTok.start, false);
		}
		if (this.isOp(p2, ":")) {
			// 3-D span `Sheet1:Sheet3!A1` — only when a sheet name and a '!' actually follow the ':'.
			const p3 = this.peek(2);
			const p4 = this.peek(3);
			if ((p3.type === "name" || p3.type === "sheet") && this.isOp(p4, "!")) {
				this.next(); // name
				this.next(); // ':'
				const to = this.next(); // second sheet name
				this.next(); // '!'
				return this.parseSheetCore(
					{ name: nameTok.value, toName: to.value },
					nameTok.start,
					false,
				);
			}
		}
		if (isBooleanLiteral(nameTok.value)) {
			this.next();
			return { type: "boolean", value: /^true$/i.test(nameTok.value) };
		}
		this.next();
		return { type: "name", name: nameTok.value };
	}

	// A quoted sheet reference: `'My Sheet'!A1`. A quoted name is external when it carries workbook or
	// path material — `'[Book.xlsx]Sheet'!A1`, `'C:\dir\f.xlsx'!A1`, `'https://h/[Book]Sheet'!A1` —
	// which stays opaque. The detection keys on `[ ] / \`: Excel forbids all four in a real sheet
	// name, so any of them marks a workbook/drive/URL, never a plain sheet. Checking them BEFORE the
	// colon split matters — a drive letter or URL scheme (`C:`, `https:`) must not be mistaken for a
	// 3-D separator. A plain quoted name with a bare colon is a 3-D span (`'Sheet 1:Sheet 3'!A1`);
	// `:` too is illegal in a sheet name, so its only meaning here is the span separator.
	private parseQuotedSheetRef(): FormulaAst {
		const sheetTok = this.next();
		this.expectOp("!");
		const external = /[[\]/\\]/.test(sheetTok.value);
		if (external) return this.parseSheetCore({ name: sheetTok.value }, sheetTok.start, true);
		const colon = sheetTok.value.indexOf(":");
		const sheet: SheetSpec =
			colon === -1
				? { name: sheetTok.value }
				: { name: sheetTok.value.slice(0, colon), toName: sheetTok.value.slice(colon + 1) };
		return this.parseSheetCore(sheet, sheetTok.start, false);
	}

	// The reference core after `sheet!`: a cell, an error literal, a defined name, or a bare
	// column/row endpoint (`Sheet1!A:A`, `Sheet1!1:1` — the `:` range is formed by the caller's loop).
	// When `external`, the whole span from `refStart` to the core is kept as one opaque node.
	private parseSheetCore(sheet: SheetSpec, refStart: number, external: boolean): FormulaAst {
		const t = this.peek();
		let end: number;
		let node: FormulaAst;
		switch (t.type) {
			case "cell":
				this.next();
				node = { type: "cell", ref: t.value, sheet };
				end = t.end;
				break;
			case "err":
				this.next();
				node = { type: "error", value: t.value };
				end = t.end;
				break;
			case "name":
			case "num":
				this.next();
				node = { type: "name", name: t.value, sheet };
				end = t.end;
				break;
			default:
				return this.fail("parse-error", "expected a reference after '!'", t);
		}
		if (external) return { type: "external", source: this.source.slice(refStart, end) };
		return node;
	}

	// An external-workbook reference introduced by a `[n]` workbook token: `[1]!Name`, `[1]Sheet1!A1`.
	// Kept opaque — we do not resolve other workbooks in v0.8.
	private parseExternalBracketRef(): FormulaAst {
		const bracket = this.next(); // [n]
		const after = this.peek();
		if (after.type === "name" || after.type === "sheet") this.next(); // optional sheet name
		this.expectOp("!");
		const core = this.peek();
		if (
			core.type !== "cell" &&
			core.type !== "name" &&
			core.type !== "err" &&
			core.type !== "num"
		) {
			return this.fail("parse-error", "expected a reference after external '!'", core);
		}
		this.next();
		return { type: "external", source: this.source.slice(bracket.start, core.end) };
	}

	private parseCall(nameTok: Token): FormulaAst {
		this.next(); // name
		this.expectOp("(");
		if (++this.funcDepth > MAX_FUNCTION_DEPTH) {
			throw new FormulaError("depth-exceeded", "function nesting too deep", {
				position: nameTok.start,
			});
		}
		const args: FormulaAst[] = [];
		if (!this.isOp(this.peek(), ")")) {
			for (;;) {
				const t = this.peek();
				// An omitted argument: `SUM(1,,2)` or a trailing `,)`.
				if (this.isOp(t, ",") || this.isOp(t, ")")) args.push({ type: "empty" });
				else args.push(this.parseExpr(0));
				if (args.length > MAX_ARGS) {
					throw new FormulaError("too-many-args", "too many function arguments", {
						position: nameTok.start,
					});
				}
				const sep = this.peek();
				if (this.isOp(sep, ",")) {
					this.next();
					continue;
				}
				if (this.isOp(sep, ")")) break;
				this.fail("parse-error", "expected ',' or ')' in argument list", sep);
			}
		}
		this.next(); // ')'
		this.funcDepth--;
		return { type: "call", name: nameTok.value, args };
	}
}

/**
 * Parse a stored-form (`ECMA-376 §18.17`) Excel formula into a typed {@link FormulaAst}. The input
 * is the formula text WITHOUT a leading `=` (what the reader's `Worksheet.formula` returns); a lone
 * leading `=` is tolerated for callers passing the display form. Parsing is structural only — names
 * are not resolved and nothing is evaluated (F8.2). Throws a typed {@link FormulaError} on a
 * malformed form, an over-deep nesting bomb, or an over-long argument list — never a bare error.
 */
export function parseFormula(text: string): FormulaAst {
	const source = text.charCodeAt(0) === 61 ? text.slice(1) : text; // strip a leading '='
	const tokens = tokenize(source);
	return new Parser(source, tokens).parse();
}
