// Public surface of the `openjsxl/formula` entry point (M8). Opt-in and module-graph-isolated: a
// consumer who never imports this path never loads a byte of the parser/evaluator. F8.1 ships the
// parser (`parseFormula`); F8.2 adds the evaluator (`evaluateWorkbook`/`evaluateCell`).

export type {
	ArrayElement,
	ArrayLiteral,
	BinaryOp,
	BooleanLiteral,
	CellRefNode,
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
// Evaluator (F8.2).
export {
	type CellResult,
	type EvaluateOptions,
	evaluateCell,
	evaluateWorkbook,
	type WorkbookEvalResult,
} from "./eval";
export type {
	ArgThunk,
	EagerFunctionSpec,
	EvalContext,
	FunctionSpec,
	LazyFunctionSpec,
} from "./functions";
export { parseFormula } from "./parser";
export {
	type ErrorCode,
	type EvalValue,
	errorValue,
	type FormulaErrorValue,
	isErrorValue,
	isRangeView,
	type RangeEntry,
	RangeView,
	type ScalarValue,
} from "./value";
