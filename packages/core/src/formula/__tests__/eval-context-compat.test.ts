import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { writeXlsx } from "../../writer";
import { BUILTIN_ENTRIES } from "../builtins";
import {
	type EagerFunctionSpec,
	type EvalContext,
	type EvalValue,
	errorValue,
	evaluateCell,
} from "../index";

// F11.33: `EvalContext.date1904` is optional, so the context shapes a 1.0.0 consumer builds by hand
// still compile. The type checker is the real assertion here: with a required member, each shape
// below is a compile error (TS2741, TS2345, TS2420, TS1360). At run time a context without the flag is
// the 1900 system, only `true` names the 1904 system, and the evaluator always sets the flag.

const now = () => 43845;
const random = () => 0.5;

// 1.0.0 shape: a literal annotated as the type.
const annotated: EvalContext = { now, random };

// 1.0.0 shape: a class that implements the interface.
class FixedContext implements EvalContext {
	now(): number {
		return 43845;
	}
	random(): number {
		return 0.5;
	}
}

// 1.0.0 shape: a literal checked with `satisfies`.
const satisfied = { now, random } satisfies EvalContext;

// A user function as a 1.0.0 consumer writes and unit-tests it.
const EPOCH: EagerFunctionSpec = {
	minArgs: 0,
	maxArgs: 0,
	evaluate: (_args, ctx) => (ctx.date1904 === true ? 1904 : 1900),
};

function invoke(name: string, args: readonly EvalValue[], ctx: EvalContext): EvalValue {
	const spec = BUILTIN_ENTRIES.find(([key]) => key === name)?.[1];
	if (spec === undefined || spec.lazyArgs) throw new Error(`missing eager builtin ${name}`);
	return spec.invoke(args, ctx);
}

// Each date built-in on one input, so the 1900 and 1904 answers sit side by side.
function dateResults(ctx: EvalContext) {
	return {
		date: invoke("DATE", [1904, 1, 1], ctx),
		year: invoke("YEAR", [0], ctx),
		weekday: invoke("WEEKDAY", [0], ctx),
		days: invoke("DAYS", [2957004, 0], ctx),
		edate: invoke("EDATE", [0, 1], ctx),
		eomonth: invoke("EOMONTH", [0, 0], ctx),
	};
}

const RESULTS_1900 = { date: 1462, year: 1899, weekday: 7, days: 2957004, edate: 31, eomonth: 1 };
const RESULTS_1904 = {
	date: 0,
	year: 1904,
	weekday: 6,
	days: errorValue("#NUM!"),
	edate: 31,
	eomonth: 30,
};

describe("an EvalContext built as in 1.0.0", () => {
	it("compiles in every shape and reads as the 1900 system", () => {
		const contexts: readonly EvalContext[] = [annotated, new FixedContext(), satisfied];
		for (const ctx of contexts) expect(dateResults(ctx)).toEqual(RESULTS_1900);
		// 1.0.0 shape: an inline literal handed straight to a function's evaluate.
		expect(EPOCH.evaluate([], { now, random })).toBe(1900);
	});

	it("is the 1900 system with the flag set false", () => {
		expect(dateResults({ now, random, date1904: false })).toEqual(RESULTS_1900);
		expect(EPOCH.evaluate([], { now, random, date1904: false })).toBe(1900);
	});

	it("still gives 1904 results with the flag set by hand", () => {
		expect(dateResults({ now, random, date1904: true })).toEqual(RESULTS_1904);
		expect(EPOCH.evaluate([], { now, random, date1904: true })).toBe(1904);
	});

	it("names the 1904 system with `true` only", () => {
		// A JS caller can pass a flag the types forbid.
		for (const flag of [1, "true", {}]) {
			expect(dateResults({ now, random, date1904: flag as never })).toEqual(RESULTS_1900);
		}
	});

	it("reads the flag once per call", () => {
		// A flag that reads false once, then true.
		let reads = 0;
		const ctx: EvalContext = {
			now,
			random,
			get date1904() {
				reads++;
				return reads > 1;
			},
		};
		for (const [name, args] of [
			["DATE", [1904, 1, 1]],
			["YEAR", [0]],
			["WEEKDAY", [0]],
			["DAYS", [2957004, 0]],
			["EDATE", [0, 1]],
			["EOMONTH", [0, 0]],
		] as const) {
			reads = 0;
			invoke(name, args, ctx);
			expect([name, reads]).toEqual([name, 1]);
		}
		// So one result never mixes the two systems: 9999-12-31 stays in range, and its year stays 9999.
		reads = 0;
		expect(invoke("DATE", [9999, 12, 31], ctx)).toBe(2958465);
		reads = 0;
		expect(invoke("YEAR", [2958465], ctx)).toBe(9999);
	});
});

describe("the evaluator's EvalContext", () => {
	it("always sets the flag, false for a 1900 workbook and true for a 1904 one", async () => {
		const seen: unknown[] = [];
		const functions = {
			FLAG: {
				minArgs: 0,
				maxArgs: 0,
				evaluate: (_args: readonly EvalValue[], ctx: EvalContext) => {
					seen.push(ctx.date1904);
					return 0;
				},
			},
		};
		for (const date1904 of [false, true]) {
			const rows = [[{ formula: "FLAG()", value: 0 }]];
			const book = await openXlsx(
				await writeXlsx({ sheets: [{ name: "S", rows }] }, { date1904 }),
			);
			await evaluateCell(book, "S", "A1", { functions });
		}
		expect(seen).toEqual([false, true]);
	});
});
