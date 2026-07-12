import { describe, expect, it } from "vitest";
import { openXlsx } from "../../reader/workbook";
import { type CellInput, writeXlsx } from "../../writer";
import { writeZip } from "../../writer/zip";
import { FormulaError } from "../errors";
import { type EvaluateOptions, evaluateCell, evaluateWorkbook } from "../eval";
import { type EvalValue, errorValue, isRangeView } from "../value";

// End-to-end evaluator tests: author a workbook (writeXlsx → openXlsx → evaluate), so the whole pipe
// is exercised on real files. The coercion RULES are unit-pinned in coerce.test.ts and were
// cross-checked cell-for-cell against Python `formulas` out-of-tree; here we pin behavior end to end
// plus the machinery the units can't reach — cycles, dependency chains, functions, defined names, and
// the adversarial-input bounds.

const f = (formula: string): CellInput => ({ formula, value: 0 });

async function evalSheet(
	rows: readonly (readonly CellInput[])[],
	options?: EvaluateOptions,
): Promise<(ref: string) => EvalValue | undefined> {
	const bytes = await writeXlsx({ sheets: [{ name: "S", rows }] });
	const book = await openXlsx(bytes);
	const result = await evaluateWorkbook(book, options);
	return (ref) => result.get("S", ref);
}

const err = (code: Parameters<typeof errorValue>[0]) => errorValue(code);

describe("evaluateWorkbook — operators & coercion end to end", () => {
	it("evaluates arithmetic, concat and comparisons over cell references", async () => {
		const g = await evalSheet([
			[10, 20, "5", true, null], // A1 B1 C1 D1 E1
			[f("A1+B1"), f("A1/B1"), f("C1+1"), f("D1+1"), f("E1+1")], // A2..E2
			[f('A1&"x"'), f("A1>B1"), f("E1=0"), f('E1=""'), f("-A1")], // A3..E3
		]);
		expect(g("A2")).toBe(30);
		expect(g("B2")).toBe(0.5);
		expect(g("C2")).toBe(6); // "5" + 1
		expect(g("D2")).toBe(2); // TRUE + 1
		expect(g("E2")).toBe(1); // empty + 1
		expect(g("A3")).toBe("10x");
		expect(g("B3")).toBe(false);
		expect(g("C3")).toBe(true); // empty = 0
		expect(g("D3")).toBe(true); // empty = ""
		expect(g("E3")).toBe(-10);
	});

	it("propagates in-sheet errors as values without aborting other cells", async () => {
		const g = await evalSheet([
			[10, 0],
			[f("A1/B1"), f("A1*2")], // A2 = #DIV/0!, B2 = 20 (unaffected)
			[f("A2+1")], // A3 propagates the error
		]);
		expect(g("A2")).toEqual(err("#DIV/0!"));
		expect(g("B2")).toBe(20);
		expect(g("A3")).toEqual(err("#DIV/0!"));
	});
});

describe("evaluateWorkbook — cycles", () => {
	it("resolves a self-reference and a 2-cell cycle to #CYCLE! without aborting", async () => {
		const g = await evalSheet([
			[f("A1"), f("C1"), f("B1")], // A1=A1 (self), B1=C1, C1=B1 (pair)
			[5, f("A1+D1"), f("100")], // A2=5 value; B2 references the cyclic A1 → propagates
		]);
		expect(g("A1")).toEqual(err("#CYCLE!"));
		expect(g("B1")).toEqual(err("#CYCLE!"));
		expect(g("C1")).toEqual(err("#CYCLE!"));
		expect(g("B2")).toEqual(err("#CYCLE!")); // depends on the cycle, but is not itself a cycle
	});
});

describe("evaluateWorkbook — function dispatch", () => {
	const functions: EvaluateOptions["functions"] = {
		MYSUM: {
			minArgs: 1,
			maxArgs: 255,
			evaluate: (args: readonly EvalValue[]) =>
				args.reduce((a: number, v) => a + (typeof v === "number" ? v : 0), 0),
		},
		MYRSUM: {
			minArgs: 1,
			maxArgs: 1,
			evaluate: (args: readonly EvalValue[]) => {
				let s = 0;
				const rv = args[0];
				if (isRangeView(rv)) for (const v of rv.values()) if (typeof v === "number") s += v;
				return s;
			},
		},
		MYIF: {
			minArgs: 3,
			maxArgs: 3,
			lazyArgs: true,
			evaluate: (thunks: readonly (() => EvalValue)[]) => {
				const c = thunks[0]?.();
				return c ? (thunks[1]?.() ?? null) : (thunks[2]?.() ?? null);
			},
		},
		MYVOL: {
			minArgs: 0,
			maxArgs: 0,
			volatile: true,
			evaluate: (_args: readonly EvalValue[], ctx: { random(): number }) => ctx.random(),
		},
	};

	it("dispatches eager, range-consuming, and lazy functions", async () => {
		const g = await evalSheet(
			[
				[1, 2, 3, 4],
				[f("MYSUM(A1,B1,C1,D1)")], // A2 = 10
				[f("MYRSUM(A1:D1)")], // A3 = 10 (RangeView)
				[f("MYIF(A1>0,A1,B1/0)")], // A4 = 1 (untaken B1/0 never evaluated)
			],
			{ functions },
		);
		expect(g("A2")).toBe(10);
		expect(g("A3")).toBe(10);
		expect(g("A4")).toBe(1); // lazy: no #DIV/0! from the untaken branch
	});

	it("returns #NAME? for an unknown function and #VALUE! for a bad arg count", async () => {
		const g = await evalSheet([[f("NOSUCH(1)"), f("MYSUM()")]], { functions });
		expect(g("A1")).toEqual(err("#NAME?"));
		expect(g("B1")).toEqual(err("#VALUE!"));
	});

	it("gates volatile functions on options.now/random (decision 3)", async () => {
		const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[f("MYVOL()")]] }] });
		const book = await openXlsx(bytes);
		await expect(evaluateCell(book, "S", "A1", { functions })).rejects.toMatchObject({
			code: "volatile-unconfigured",
		});
		const injected = await evaluateCell(book, "S", "A1", { functions, random: () => 0.42 });
		expect(injected).toBe(0.42);
	});
});

// A minimal hand-crafted xlsx: the writer can't author <definedNames>, so we build the parts directly.
const enc = new TextEncoder();
function craftWorkbook(sheetData: string, definedNames = ""): Promise<Uint8Array> {
	const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	return writeZip([
		{
			name: "[Content_Types].xml",
			data: enc.encode(
				'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
					'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
					'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
					'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
			),
		},
		{
			name: "_rels/.rels",
			data: enc.encode(
				`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rns}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
			),
		},
		{
			name: "xl/workbook.xml",
			data: enc.encode(
				`<?xml version="1.0"?><workbook xmlns="${ns}" xmlns:r="${rns}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>${definedNames}</workbook>`,
			),
		},
		{
			name: "xl/_rels/workbook.xml.rels",
			data: enc.encode(
				`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rns}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
			),
		},
		{
			name: "xl/worksheets/sheet1.xml",
			data: enc.encode(
				`<?xml version="1.0"?><worksheet xmlns="${ns}"><sheetData>${sheetData}</sheetData></worksheet>`,
			),
		},
	]);
}

describe("evaluateWorkbook — defined names", () => {
	it("resolves a constant, a single-cell ref, and a range name; #NAME? otherwise", async () => {
		const sheetData =
			'<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>100</v></c><c r="C1"><f>TAX*2</f><v>0</v></c></row>' +
			'<row r="2"><c r="A2"><v>20</v></c><c r="C2"><f>RATE*100</f><v>0</v></c></row>' +
			'<row r="3"><c r="A3"><v>30</v></c><c r="C3"><f>MYRSUM(MYRANGE)</f><v>0</v></c><c r="D3"><f>NOSUCHNAME+1</f><v>0</v></c></row>';
		const definedNames =
			"<definedNames>" +
			'<definedName name="TAX">Data!$B$1</definedName>' +
			'<definedName name="RATE">0.1</definedName>' +
			'<definedName name="MYRANGE">Data!$A$1:$A$3</definedName>' +
			"</definedNames>";
		const book = await openXlsx(await craftWorkbook(sheetData, definedNames));
		expect(book.definedNames.map((d) => d.name)).toEqual(["TAX", "RATE", "MYRANGE"]);
		const functions: EvaluateOptions["functions"] = {
			MYRSUM: {
				minArgs: 1,
				maxArgs: 1,
				evaluate: (args: readonly EvalValue[]) => {
					let s = 0;
					const rv = args[0];
					if (isRangeView(rv))
						for (const v of rv.values()) if (typeof v === "number") s += v;
					return s;
				},
			},
		};
		const r = await evaluateWorkbook(book, { functions });
		expect(r.get("Data", "C1")).toBe(200); // TAX (Data!$B$1 = 100) * 2
		expect(r.get("Data", "C2")).toBe(10); // RATE (0.1) * 100
		expect(r.get("Data", "C3")).toBe(60); // MYRSUM over MYRANGE (10+20+30)
		expect(r.get("Data", "D3")).toEqual(err("#NAME?")); // unknown name
	});
});

describe("evaluateWorkbook — adversarial input is bounded and typed", () => {
	it("evaluates a deep reference chain without growing the native stack", async () => {
		// 30k is well past the ~10k native-recursion limit a naive evaluator would blow.
		const n = 30000;
		const rows: CellInput[][] = [[1]];
		for (let i = 2; i <= n; i++) rows.push([f(`A${i - 1}+1`)]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		expect(await evaluateCell(book, "S", `A${n}`)).toBe(n);
	});

	it("evaluates a whole-grid range in bounded time (only used cells iterated)", async () => {
		const functions: EvaluateOptions["functions"] = {
			MYRSUM: {
				minArgs: 1,
				maxArgs: 1,
				evaluate: (args: readonly EvalValue[]) => {
					let s = 0;
					const rv = args[0];
					if (isRangeView(rv))
						for (const v of rv.values()) if (typeof v === "number") s += v;
					return s;
				},
			},
		};
		const g = await evalSheet([[1], [2], [3], [], [f("MYRSUM(A1:XFD1048576)")]], { functions });
		expect(g("A5")).toBe(6); // 1+2+3 (A5 itself is a formula → 0 in the sum)
	});

	it("aborts with a typed budget error when the cell-visit fuel runs out", async () => {
		const rows: CellInput[][] = [[1]];
		for (let i = 2; i <= 100; i++) rows.push([f(`A${i - 1}+1`)]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		await expect(evaluateCell(book, "S", "A100", { maxCellVisits: 10 })).rejects.toBeInstanceOf(
			FormulaError,
		);
	});
});

// Regressions pinned from the F8.2 adversarial review. Each was an unbounded hang or an UNTYPED
// native failure before the fix.
describe("evaluateWorkbook — adversarial-review regressions (F8.2)", () => {
	it("does not loop on a self-referential range formula (A1 = A1:A1)", async () => {
		// Range construction is lazy, so tri-color cycle detection can't see the self-reference; the
		// unwrap guard in reduce/scalarize makes it a typed error VALUE instead of an infinite loop.
		const single = await evalSheet([[f("A1:A1")]]);
		expect(single("A1")).toEqual(err("#CYCLE!"));
		const multi = await evalSheet([[f("A1:B2")]]);
		expect(multi("A1")).toEqual(err("#CYCLE!"));
	});

	it("caps a deep left-associative operator chain with a typed error, not a stack overflow", async () => {
		// `1+1+…` is left-associative, which the parser builds ITERATIVELY (escaping its own depth
		// cap), so the AST is as deep as the term count; the evaluator's native-recursion guard turns
		// a would-be RangeError into a typed FormulaError. A shallow chain still evaluates.
		const shallow = await evalSheet([[f(Array.from({ length: 100 }, () => "1").join("+"))]]);
		expect(shallow("A1")).toBe(100);
		const deep = Array.from({ length: 2000 }, () => "1").join("+"); // ~3999 chars, under the writer cap
		const book = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [[f(deep)]] }] }),
		);
		await expect(evaluateCell(book, "S", "A1")).rejects.toMatchObject({
			code: "depth-exceeded",
		});
	});

	it("caps deep nesting through lazy-function args with a typed error, not a stack overflow", async () => {
		const functions: EvaluateOptions["functions"] = {
			LZ: {
				minArgs: 1,
				maxArgs: 1,
				lazyArgs: true,
				evaluate: (thunks: readonly (() => EvalValue)[]) => thunks[0]?.() ?? null,
			},
		};
		const rows: CellInput[][] = [[1]];
		for (let i = 2; i <= 5000; i++) rows.push([f(`LZ(A${i - 1})`)]);
		const book = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }));
		await expect(evaluateCell(book, "S", "A5000", { functions })).rejects.toBeInstanceOf(
			FormulaError,
		);
	});
});
