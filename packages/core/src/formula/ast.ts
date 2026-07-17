// The typed AST that `parseFormula` produces from a stored-form (`ECMA-376 §18.17`) formula string.
// F8.1 is parse-only: these nodes describe structure, not results — evaluation (F8.2) walks them.
// The tree is faithful to the surface syntax; it does NOT resolve names, coerce types, or decide
// what a reference denotes. Two deliberate opacity escapes — `StructuredRef` and `ExternalRef` —
// keep constructs we do not yet interpret as verbatim source spans so nothing is silently mangled.

/** A qualifying sheet (or 3-D sheet span) on a reference: `Sheet1!`, `Sheet1:Sheet3!`. */
export interface SheetSpec {
	/** The (unescaped) first sheet name. */
	readonly name: string;
	/** The end sheet of a 3-D span (`Sheet1:Sheet3!A1`); absent for a single-sheet reference. */
	readonly toName?: string;
}

/** A numeric literal (`3`, `1.5`, `1E+20`). `raw` preserves the exact token as written. */
export interface NumberLiteral {
	readonly type: "number";
	readonly value: number;
	readonly raw: string;
}

/** A string literal; `value` is already unescaped (`""` → `"`). */
export interface StringLiteral {
	readonly type: "string";
	readonly value: string;
}

/** A bare boolean literal (`TRUE`/`FALSE`). `TRUE()`/`FALSE()` are `FunctionCall`s instead. */
export interface BooleanLiteral {
	readonly type: "boolean";
	readonly value: boolean;
}

/** A cell error literal — one of the eight `ST_CellErrorType` values, e.g. `"#DIV/0!"`. */
export interface ErrorLiteral {
	readonly type: "error";
	readonly value: string;
}

/** An array constant `{1,2;3,4}`: `rows[r][c]`, `,` = columns, `;` = rows. Constants only. */
export interface ArrayLiteral {
	readonly type: "array";
	readonly rows: readonly (readonly ArrayElement[])[];
}

/** An element inside an array constant: a literal, optionally negated/plussed. */
export type ArrayElement = NumberLiteral | StringLiteral | BooleanLiteral | ErrorLiteral | UnaryOp;

/**
 * A single cell reference such as `A1`, `$A$1`, `Sheet1!B2` — `ref` keeps the `$` markers.
 * Named `CellRefNode` (not `CellRef`) so the `openjsxl/formula` surface never collides with the
 * `CellRef` (`{col,row}`) that `openjsxl` exports — a consumer can import both entry points at once.
 */
export interface CellRefNode {
	readonly type: "cell";
	readonly ref: string;
	readonly sheet?: SheetSpec;
}

/** A bare name node: a defined name, or a whole-column/row endpoint (`A` in `A:A`, `1` in `1:1`). */
export interface NameRef {
	readonly type: "name";
	readonly name: string;
	readonly sheet?: SheetSpec;
}

/** The range operator `:` — `A1:B2`, `A:A`, `1:1`, or `INDEX(...):INDEX(...)`. */
export interface RangeRef {
	readonly type: "range";
	readonly left: FormulaAst;
	readonly right: FormulaAst;
}

/** The reference union operator `,` inside grouping parens: `(A1,A3,B5)`. */
export interface UnionRef {
	readonly type: "union";
	readonly items: readonly FormulaAst[];
}

/** A function call `NAME(arg, …)`. An omitted argument (`SUM(A1,,A2)`) is an `EmptyArg`. */
export interface FunctionCall {
	readonly type: "call";
	readonly name: string;
	readonly args: readonly FormulaAst[];
}

/** An omitted call argument — the gap in `SUM(1,,2)`. Never appears outside an argument list. */
export interface EmptyArg {
	readonly type: "empty";
}

/** A prefix operator: unary minus/plus (`-A1`) or implicit intersection (`@A1:B2`). */
export interface UnaryOp {
	readonly type: "unary";
	readonly op: "-" | "+" | "@";
	readonly operand: FormulaAst;
}

/** A postfix operator: percent (`50%`) or spilled-range (`A1#`). */
export interface PostfixOp {
	readonly type: "postfix";
	readonly op: "%" | "#";
	readonly operand: FormulaAst;
}

/** An infix operator: arithmetic (`+ - * / ^`), concatenation (`&`), or a comparison. */
export interface BinaryOp {
	readonly type: "binary";
	readonly op: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";
	readonly left: FormulaAst;
	readonly right: FormulaAst;
}

/** A parenthesized grouping `( … )`, kept as a node so precedence is explicit and reprintable. */
export interface Group {
	readonly type: "group";
	readonly expr: FormulaAst;
}

/** A structured (table) reference such as `Table1[@Amount]` — kept opaque as its source span. */
export interface StructuredRef {
	readonly type: "structured";
	readonly source: string;
}

/** An external-workbook reference such as `[1]!Name` or `'[Book.xlsx]Sheet'!A1` — kept opaque. */
export interface ExternalRef {
	readonly type: "external";
	readonly source: string;
}

/** Every node the parser can produce. */
export type FormulaAst =
	| NumberLiteral
	| StringLiteral
	| BooleanLiteral
	| ErrorLiteral
	| ArrayLiteral
	| CellRefNode
	| NameRef
	| RangeRef
	| UnionRef
	| FunctionCall
	| EmptyArg
	| UnaryOp
	| PostfixOp
	| BinaryOp
	| Group
	| StructuredRef
	| ExternalRef;
