import * as main from "openjsxl";
import * as formula from "openjsxl/formula";
import { describe, expect, it } from "vitest";

// Compile-time proof that `openjsxl` and `openjsxl/formula` share no exported name. A consumer that
// imports from both entry points in one module must never hit a duplicate identifier, and any name
// exported by BOTH would make the star re-export below emit TS2308 ("Module has already exported a
// member named 'X'") — failing `tsc --noEmit`, a release gate. Pre-F10.6 this did NOT compile: both
// entries exported a `CellRef` (core's `{col,row}` addressing type vs the formula AST node, now
// `CellRefNode`). This file freezes that collision-freedom for 1.0 and forever after.
export * from "openjsxl";
export * from "openjsxl/formula";

describe("entry-point export surfaces", () => {
	it("openjsxl and openjsxl/formula share no runtime value name", () => {
		const shared = Object.keys(main).filter((k) => k in formula);
		expect(shared).toEqual([]);
	});
});
