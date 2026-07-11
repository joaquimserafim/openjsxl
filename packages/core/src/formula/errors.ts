// The formula entry point's OWN error class. Kept deliberately separate from core's `XlsxError`
// so the reader/writer's closed `XlsxErrorCode` union stays frozen (M8 decision 4): a caller who
// never imports `openjsxl/formula` never sees these codes, and adding a formula failure mode here
// can never widen the file-level error surface.
//
// These are CONFIGURATION / STRUCTURAL failures of parsing or (later) evaluation — a malformed
// stored form, a nesting bomb, an argument-count overflow. They are distinct from *cell* error
// values (`#REF!`, `#VALUE!`, …), which are ordinary data that parse into `ErrorLiteral` AST nodes
// and propagate as values; a `#DIV/0!` in a formula is not a `FormulaError`.

export type FormulaErrorCode =
	| "parse-error" // the text is not a well-formed stored-form formula
	| "depth-exceeded" // nesting (parens / unary chain / calls) passed the safety cap
	| "too-many-args" // a function call carried more than the 255-argument maximum
	| "unsupported"; // a construct we recognize but do not yet handle at this layer

export class FormulaError extends Error {
	/** Machine-readable discriminant; branch on this rather than the message. */
	readonly code: FormulaErrorCode;
	/** 0-based character offset into the formula text where the failure was detected, when known. */
	readonly position?: number;

	constructor(
		code: FormulaErrorCode,
		message: string,
		options?: { cause?: unknown; position?: number },
	) {
		super(message, options);
		this.name = "FormulaError";
		this.code = code;
		if (options?.position !== undefined) this.position = options.position;
	}
}
