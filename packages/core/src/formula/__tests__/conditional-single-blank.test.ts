import { describe, expect, it, vi } from "vitest";
import { MAX_ROW } from "../../ooxml/a1";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import {
	BUILTIN_ENTRIES,
	type Criteria,
	conditionalAggregate,
	parseCriteria,
	valueRangeAggregate,
} from "../builtins";
import { evaluateCell, evaluateWorkbook } from "../eval";
import type { EvalContext } from "../functions";
import { type EvalValue, errorValue, isErrorValue, RangeView } from "../value";

// F11.39: SUMIF and AVERAGEIF skipped a matched error at a blank criteria cell, and evaluated cells
// SUMIFS never reads, so they disagreed with SUMIFS and AVERAGEIFS (F11.12). Owner rulings (2026-11-01,
// "Walk like SUMIFS", and D20): given a sum argument, SUMIF and AVERAGEIF walk it as SUMIFS does, for
// every criterion, so the two families agree on every value and error and resolve the same cells in
// the same order.

const NA = { formula: "NA()", value: 0 };
const DIV0 = { formula: "1/0", value: 0 };
const NOW = { formula: "NOW()", value: 0 };
const EMPTY_TEXT = { formula: '""', value: "" };

async function book(rows: readonly (readonly CellInput[])[], formulas: readonly string[]) {
	return openXlsx(
		await writeXlsx({
			sheets: [
				{ name: "Data", rows },
				{ name: "Results", rows: formulas.map((formula) => [{ formula, value: 0 }]) },
			],
		}),
	);
}

// Every formula's value, through evaluateWorkbook (which configures no clock).
async function evaluate(rows: readonly (readonly CellInput[])[], formulas: readonly string[]) {
	const result = await evaluateWorkbook(await book(rows, formulas));
	return formulas.map((_, index) => result.get("Results", `A${index + 1}`));
}

// Each formula through evaluateCell alone, so a volatile cell it resolves throws.
async function evaluateEach(rows: readonly (readonly CellInput[])[], formulas: readonly string[]) {
	const workbook = await book(rows, formulas);
	const values: unknown[] = [];
	for (let i = 0; i < formulas.length; i++) {
		values.push(await evaluateCell(workbook, "Results", `A${i + 1}`));
	}
	return values;
}

describe("SUMIF and AVERAGEIF agree with SUMIFS and AVERAGEIFS on blank criteria cells (F11.39)", () => {
	it("propagates a matched error at a blank criteria cell", async () => {
		const values = await evaluate(
			[
				[1, 1],
				[null, NA],
			],
			[
				'SUMIF(Data!A1:A2,"",Data!B1:B2)',
				'SUMIFS(Data!B1:B2,Data!A1:A2,"")',
				'AVERAGEIF(Data!A1:A2,"",Data!B1:B2)',
				'AVERAGEIFS(Data!B1:B2,Data!A1:A2,"")',
			],
		);
		expect(values).toEqual(Array(4).fill(errorValue("#N/A")));
	});

	it("propagates a matched error under a not-equal criterion", async () => {
		const values = await evaluate(
			[
				[null, 5, "a"],
				[null, NA, null],
			],
			['SUMIF(Data!C1:C2,"<>q",Data!B1:B2)', 'SUMIFS(Data!B1:B2,Data!C1:C2,"<>q")'],
		);
		expect(values).toEqual([errorValue("#N/A"), errorValue("#N/A")]);
	});

	it("still propagates a matched error at a criteria formula returning empty text", async () => {
		const values = await evaluate(
			[
				[1, 1],
				[EMPTY_TEXT, NA],
			],
			['SUMIF(Data!A1:A2,"",Data!B1:B2)', 'SUMIFS(Data!B1:B2,Data!A1:A2,"")'],
		);
		expect(values).toEqual([errorValue("#N/A"), errorValue("#N/A")]);
	});

	it("returns the first matched error in sheet order, as SUMIFS does", async () => {
		for (const rows of [
			[
				[null, NA],
				[EMPTY_TEXT, DIV0],
			],
			[
				[EMPTY_TEXT, DIV0],
				[null, NA],
			],
		]) {
			const [sumif, sumifs, averageif, averageifs] = await evaluate(rows, [
				'SUMIF(Data!A1:A2,"",Data!B1:B2)',
				'SUMIFS(Data!B1:B2,Data!A1:A2,"")',
				'AVERAGEIF(Data!A1:A2,"",Data!B1:B2)',
				'AVERAGEIFS(Data!B1:B2,Data!A1:A2,"")',
			]);
			expect(sumif).toEqual(sumifs);
			expect(averageif).toEqual(averageifs);
		}
		const [first] = await evaluate(
			[
				[null, NA],
				[EMPTY_TEXT, DIV0],
			],
			['SUMIF(Data!A1:A2,"",Data!B1:B2)'],
		);
		expect(first).toEqual(errorValue("#N/A"));
	});

	it("never evaluates a value the criterion excludes", async () => {
		const values = await evaluateEach(
			[
				["x", NOW],
				[null, 5],
			],
			[
				'SUMIF(Data!A1:A2,"<>x",Data!B1:B2)',
				'SUMIFS(Data!B1:B2,Data!A1:A2,"<>x")',
				'AVERAGEIF(Data!A1:A2,"<>x",Data!B1:B2)',
				'AVERAGEIFS(Data!B1:B2,Data!A1:A2,"<>x")',
			],
		);
		// No clock is configured, so evaluating NOW() would throw volatile-unconfigured.
		expect(values).toEqual([5, 5, 5, 5]);
	});

	it("stops at the first matched error and never reaches a later volatile value", async () => {
		const values = await evaluateEach(
			[
				[null, NA],
				[EMPTY_TEXT, NOW],
			],
			['SUMIF(Data!A1:A2,"",Data!B1:B2)', 'SUMIFS(Data!B1:B2,Data!A1:A2,"")'],
		);
		expect(values).toEqual([errorValue("#N/A"), errorValue("#N/A")]);
	});

	it("never evaluates a criteria cell whose value cell is blank", async () => {
		const values = await evaluateEach(
			[
				[NOW, null],
				[null, 5],
			],
			['SUMIF(Data!A1:A2,"<>x",Data!B1:B2)', 'SUMIFS(Data!B1:B2,Data!A1:A2,"<>x")'],
		);
		expect(values).toEqual([5, 5]);
	});

	it("never evaluates such a criteria cell under a criterion a blank cannot match (D20)", async () => {
		const values = await evaluateEach(
			[
				[NOW, null],
				[1, 5],
			],
			['SUMIF(Data!A1:A2,">0",Data!B1:B2)', 'SUMIFS(Data!B1:B2,Data!A1:A2,">0")'],
		);
		expect(values).toEqual([5, 5]);
	});

	it("walks a single-cell sum argument too (D20)", async () => {
		const values = await evaluate([[null, 5]], ['SUMIF(Data!A1:A1,"",Data!B1)']);
		expect(values).toEqual([5]);
	});

	it("walks a blank single-cell sum argument, never summing the criteria cells", async () => {
		const values = await evaluate(
			[[5, null], [7]],
			[
				'SUMIF(Data!A1,">0",Data!B1)',
				'SUMIFS(Data!B1,Data!A1,">0")',
				'SUMIF(Data!A1:A2,">0",Data!B1)',
				'AVERAGEIF(Data!A1,">0",Data!B1)',
				'AVERAGEIFS(Data!B1,Data!A1,">0")',
			],
		);
		expect(values).toEqual([0, 0, 0, errorValue("#DIV/0!"), errorValue("#DIV/0!")]);
	});
});

// ── the same functions on hand-built ranges ──────────────────────────────────────────────────────

const context: EvalContext = { date1904: false, now: () => 0, random: () => 0 };

function invoke(name: string, args: readonly EvalValue[]): EvalValue {
	const spec = BUILTIN_ENTRIES.find(([key]) => key === name)?.[1];
	if (spec === undefined || spec.lazyArgs) throw new Error(`missing eager builtin ${name}`);
	return spec.invoke(args, context);
}

const BLANK = Symbol("blank");
type Slot = EvalValue | typeof BLANK;

// A one-column range on sheet S whose `slots` fill rows 1.., a BLANK slot unpopulated. Every cell it
// resolves is appended to `log` as `${tag}${row}`.
function column(tag: string, col: number, slots: readonly Slot[], log: string[]): RangeView {
	const populated: [number, number][] = [];
	slots.forEach((slot, i) => {
		if (slot !== BLANK) populated.push([col, i + 1]);
	});
	return new RangeView(
		"S",
		col,
		1,
		col,
		slots.length,
		() => populated,
		(_, row) => {
			log.push(`${tag}${row}`);
			const slot = slots[row - 1];
			return slot === undefined || slot === BLANK ? null : slot;
		},
	);
}

// Every sequence of `length` slots drawn from `choices`.
function sequences(choices: readonly Slot[], length: number): Slot[][] {
	if (length === 0) return [[]];
	return sequences(choices, length - 1).flatMap((rest) =>
		choices.map((c): Slot[] => [c, ...rest]),
	);
}

// A slot list as text, for a mismatch message: a blank slot prints as `_`.
const label = (slots: readonly Slot[]) =>
	slots.map((slot) => (slot === BLANK ? "_" : JSON.stringify(slot))).join(" ");

const CRITERIA_SLOTS: readonly Slot[] = [BLANK, "", "x", 5, errorValue("#N/A")];
const VALUE_SLOTS: readonly Slot[] = [BLANK, 3, "t", errorValue("#N/A"), errorValue("#DIV/0!")];

const PAIRS: readonly (readonly [string, string])[] = [
	["SUMIF", "SUMIFS"],
	["AVERAGEIF", "AVERAGEIFS"],
];

// Each single-criterion function and its IFS form on the same input: results and resolution logs.
function compare(criterion: string, criteria: readonly Slot[], values: readonly Slot[]) {
	const results: {
		single: EvalValue;
		multi: EvalValue;
		singleLog: string[];
		multiLog: string[];
	}[] = [];
	for (const [single, multi] of PAIRS) {
		const singleLog: string[] = [];
		const multiLog: string[] = [];
		const one = invoke(single, [
			column("A", 1, criteria, singleLog),
			criterion,
			column("B", 2, values, singleLog),
		]);
		const many = invoke(multi, [
			column("B", 2, values, multiLog),
			column("A", 1, criteria, multiLog),
			criterion,
		]);
		results.push({ single: one, multi: many, singleLog, multiLog });
	}
	return results;
}

describe("SUMIF and AVERAGEIF match the IFS forms on every three-cell input (F11.39)", () => {
	const criteriaSets = sequences(CRITERIA_SLOTS, 3);
	const valueSets = sequences(VALUE_SLOTS, 3);

	// Criteria a blank cell matches, then criteria it cannot (a number criterion matches numbers only,
	// so "<>5" cannot match a blank).
	it.each(["", "=", "<>x", "<>*", "x", ">1", "<>", "5", "<>5"])(
		"returns the same value or error and resolves the same cells in order for %j",
		(criterion) => {
			const mismatches: string[] = [];
			for (const criteria of criteriaSets) {
				for (const values of valueSets) {
					for (const r of compare(criterion, criteria, values)) {
						const same =
							JSON.stringify(r.single) === JSON.stringify(r.multi) &&
							r.singleLog.join() === r.multiLog.join();
						if (!same) mismatches.push(`${label(criteria)} / ${label(values)}`);
					}
				}
			}
			expect(mismatches.length, mismatches[0]).toBe(0);
		},
	);

	it("resolves only the populated value cells of whole columns, and their criteria cells", () => {
		const coordinates: [number, number][] = [
			[2, 10],
			[2, MAX_ROW],
		];
		const critPopulated = vi.fn((): [number, number][] => [[1, 10]]);
		const resolved: string[] = [];
		const criteria = new RangeView("S", 1, 1, 1, MAX_ROW, critPopulated, (_, row) => {
			resolved.push(`A${row}`);
			return row === 10 ? "x" : null;
		});
		const values = new RangeView(
			"S",
			2,
			1,
			2,
			MAX_ROW,
			() => coordinates,
			(_, row) => {
				resolved.push(`B${row}`);
				return row === 10 ? 1 : 4;
			},
		);
		expect(invoke("SUMIF", [criteria, "<>x", values])).toBe(4);
		expect(resolved).toEqual([`A10`, `A${MAX_ROW}`, `B${MAX_ROW}`]);
		expect(critPopulated).not.toHaveBeenCalled();
	});
});

describe("conditionalAggregate", () => {
	const log: string[] = [];
	const a = (...slots: Slot[]) => column("A", 1, slots, log);
	const b = (...slots: Slot[]) => column("B", 2, slots, log);

	it("returns a criterion error", () => {
		expect(conditionalAggregate(a("x"), errorValue("#N/A"), b(1), "sum")).toEqual(
			errorValue("#N/A"),
		);
	});

	it("walks a value range for every criterion", () => {
		expect(conditionalAggregate(a(1, BLANK), "", b(1, errorValue("#N/A")), "sum")).toEqual(
			errorValue("#N/A"),
		);
		expect(conditionalAggregate(a("x", BLANK), "<>x", b(3, 5), "avg")).toBe(5);
		expect(
			conditionalAggregate(
				a("x", "x"),
				"x",
				b(errorValue("#DIV/0!"), errorValue("#N/A")),
				"sum",
			),
		).toEqual(errorValue("#DIV/0!"));
	});

	it("walks the value range in its order for a criterion a blank cannot match too (D20)", () => {
		const order: string[] = [];
		const criteria = column("A", 1, ["y", "x"], order);
		expect(conditionalAggregate(criteria, "x", column("B", 2, [BLANK, 3], order), "sum")).toBe(
			3,
		);
		expect(order).toEqual(["A2", "B2"]);
	});

	it("counts matching cells, and the blank positions when a blank matches", () => {
		expect(conditionalAggregate(a("x", BLANK, BLANK), "", undefined, "count")).toBe(2);
		expect(conditionalAggregate(a("x", BLANK, "x"), "x", undefined, "count")).toBe(2);
	});

	it("averages a criteria range's matched numbers by itself, or is #DIV/0! with none", () => {
		expect(conditionalAggregate(a(1, 2), ">5", undefined, "avg")).toEqual(
			errorValue("#DIV/0!"),
		);
		expect(conditionalAggregate(a(1, 2), ">0", undefined, "avg")).toBe(1.5);
		expect(conditionalAggregate(a(2, "t", 4), "<>x", undefined, "avg")).toBe(3); // text not counted
		expect(conditionalAggregate(a(true, 2), "<>x", undefined, "avg")).toBe(2); // nor a Boolean
	});

	it("sums a criteria range's matched numbers by itself, reading each cell once", () => {
		expect(conditionalAggregate(a(2, BLANK, 4), "<>x", undefined, "sum")).toBe(6);
		expect(conditionalAggregate(a(true, 2), true, undefined, "sum")).toBe(0);
		expect(conditionalAggregate(a(5), ">0", null, "sum")).toBe(0); // a blank one-cell sum argument
		const order: string[] = [];
		const own = column("A", 1, [2, "t", 4], order);
		expect(conditionalAggregate(own, "<>x", undefined, "sum")).toBe(6);
		expect(order).toEqual(["A1", "A2", "A3"]);
	});
});

describe("valueRangeAggregate", () => {
	const crit = (raw: string): Criteria => {
		const parsed = parseCriteria(raw);
		if (isErrorValue(parsed)) throw new Error(`bad criterion ${raw}`);
		return parsed;
	};

	it("returns the first matched error, without resolving later values", () => {
		const log: string[] = [];
		const result = valueRangeAggregate(
			column("A", 1, [BLANK, BLANK], log),
			crit(""),
			column("B", 2, [errorValue("#N/A"), 9], log),
			"sum",
		);
		expect(result).toEqual(errorValue("#N/A"));
		expect(log).toEqual(["A1", "B1"]);
	});

	it("is #DIV/0! for an average with no number matched", () => {
		const log: string[] = [];
		expect(
			valueRangeAggregate(
				column("A", 1, [BLANK, BLANK], log),
				crit("<>x"),
				column("B", 2, [BLANK, "t"], log),
				"avg",
			),
		).toEqual(errorValue("#DIV/0!"));
	});

	it("averages the matched numbers", () => {
		const log: string[] = [];
		expect(
			valueRangeAggregate(
				column("A", 1, [BLANK, "y"], log),
				crit("<>x"),
				column("B", 2, [3, 5], log),
				"avg",
			),
		).toBe(4);
	});

	it("sums matched numbers inside the criteria rectangle, never resolving an excluded value", () => {
		const log: string[] = [];
		const result = valueRangeAggregate(
			column("A", 1, ["x", BLANK], log),
			crit("<>x"),
			column("B", 2, [errorValue("#N/A"), 5, 7], log),
			"sum",
		);
		expect(result).toBe(5);
		expect(log).toEqual(["A1", "A2", "B2"]); // B1 excluded, B3 outside A1:A2
	});

	it("skips value cells right of the criteria rectangle", () => {
		const log: string[] = [];
		const wide = new RangeView(
			"S",
			2,
			1,
			3,
			1,
			() => [
				[2, 1],
				[3, 1],
			],
			(col) => {
				log.push(`col ${col}`);
				return col === 2 ? 5 : 7;
			},
		);
		expect(valueRangeAggregate(column("A", 1, [BLANK], log), crit(""), wide, "sum")).toBe(5);
		expect(log).toEqual(["A1", "col 2"]);
	});

	it("walks a scalar value as one cell at the top-left", () => {
		const log: string[] = [];
		expect(valueRangeAggregate(column("A", 1, [BLANK, "x"], log), crit(""), 5, "sum")).toBe(5);
		expect(log).toEqual(["A1"]);
		expect(valueRangeAggregate(column("A", 1, ["x"], []), crit("<>x"), 5, "sum")).toBe(0);
	});

	it("tests a scalar criterion as its own value", () => {
		expect(valueRangeAggregate("x", crit("x"), 5, "sum")).toBe(5);
		expect(valueRangeAggregate("y", crit("x"), 5, "sum")).toBe(0);
	});

	it("reshapes to a scalar criterion's single cell", () => {
		const log: string[] = [];
		const values = column("B", 2, [3, 4], log);
		expect(valueRangeAggregate(null, crit(""), values, "sum")).toBe(3);
		expect(log).toEqual(["B1"]); // B2 is outside the one-cell criteria rectangle
	});
});
