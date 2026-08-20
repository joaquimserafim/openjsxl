import type { DecodeContext } from "../ooxml";
import type { Row } from "../types";
import type { XmlToken } from "../xml";
import { createRowAssemblyState, flushAssembler, pushToken } from "./row-assembler-state";

export interface RowAssembler {
	/** Advance the state machine by one token; returns rows completed by that token. */
	push(token: XmlToken): Row[];
	/** Emit a row left open at end of input, as can happen with truncated XML. */
	flush(): Row[];
}

/**
 * Incrementally assemble worksheet tokens into rows for buffered and streaming readers.
 *
 * Input: `<row r="5"><c><v>1</v></c><c r="D5"><v>4</v></c></row>`
 * Output: row 5 containing `A5 = 1` and `D5 = 4`; a following no-`r` cell becomes `E5`.
 * `flush()` preserves an otherwise readable final row when the XML ends early.
 */
export function createRowAssembler(ctx: DecodeContext): RowAssembler {
	const state = createRowAssemblyState(ctx);
	return {
		push: (token) => pushToken(state, token),
		flush: () => flushAssembler(state),
	};
}
