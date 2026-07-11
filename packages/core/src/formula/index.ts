// Public surface of the `openjsxl/formula` entry point (M8). Opt-in and module-graph-isolated: a
// consumer who never imports this path never loads a byte of the parser/evaluator. F8.1 ships the
// parser; F8.2+ will add evaluation behind this same entry.

export type {
	ArrayElement,
	ArrayLiteral,
	BinaryOp,
	BooleanLiteral,
	CellRef,
	EmptyArg,
	ErrorLiteral,
	ExternalRef,
	FormulaAst,
	FunctionCall,
	Group,
	NameRef,
	NumberLiteral,
	PostfixOp,
	RangeRef,
	SheetSpec,
	StringLiteral,
	StructuredRef,
	UnaryOp,
	UnionRef,
} from "./ast";
export { FormulaError, type FormulaErrorCode } from "./errors";
export { parseFormula } from "./parser";
