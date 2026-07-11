import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { workbookToInput } from "../from-workbook";
import { writeXlsx } from "../workbook";

// F5.4 — writing formula text. A `{ formula, value?, style? }` cell emits <f>…</f> with the optional
// cached result; the reader reads both back. Validation mirrors the other writer inputs: single-read,
// non-empty, XML-safe, stored form (no leading =), within Excel's length ceiling.

const one = (cell: unknown) => ({
	// biome-ignore lint/suspicious/noExplicitAny: exercising the cell input union directly
	sheets: [{ name: "S", rows: [[cell as any]] }],
});

async function roundtrip(cell: unknown) {
	const wb = await openXlsx(await writeXlsx(one(cell)));
	return wb.sheet("S");
}

describe("writeXlsx — formula cells", () => {
	it("emits a formula with its cached numeric result", async () => {
		const s = await roundtrip({ formula: "B1*2", value: 84 });
		expect(s.formula("A1")).toBe("B1*2");
		expect(s.cell("A1").value).toBe(84);
	});

	it("emits a string result as t=str, a boolean as t=b, and a date as a serial", async () => {
		expect((await roundtrip({ formula: 'B1&"!"', value: "hi!" })).cell("A1").value).toBe("hi!");
		expect((await roundtrip({ formula: "B1>0", value: true })).cell("A1").value).toBe(true);
		const d = new Date(Date.UTC(2021, 5, 15));
		const cell = (await roundtrip({ formula: "C1+1", value: d })).cell("A1");
		expect(cell.type).toBe("date");
		expect((cell.value as Date).getTime()).toBe(d.getTime());
	});

	it("emits a formula with no cached value (Excel computes on open)", async () => {
		const s = await roundtrip({ formula: "A1+1" });
		expect(s.formula("A1")).toBe("A1+1");
		expect(s.cell("A1").type).toBe("empty");
	});

	it("carries a style alongside the formula", async () => {
		const style = { font: { bold: true } } as const;
		const s = await roundtrip({ formula: "SUM(B1:B3)", value: 6, style });
		expect(s.formula("A1")).toBe("SUM(B1:B3)");
		expect(s.style("A1")).toEqual(style);
	});

	it("escapes XML metacharacters in the formula text", async () => {
		const s = await roundtrip({ formula: "A1<B1", value: true });
		expect(s.formula("A1")).toBe("A1<B1"); // round-trips through &lt;
	});
});

describe("writeXlsx — formula validation", () => {
	const reject = async (cell: unknown, pattern: RegExp): Promise<void> => {
		const err = await writeXlsx(one(cell)).then(
			() => undefined,
			(e) => e,
		);
		expect(err).toBeInstanceOf(XlsxError);
		expect((err as XlsxError).code).toBe("invalid-input");
		expect((err as XlsxError).message).toMatch(pattern);
	};

	it("rejects a leading =, empty, over-long, non-string, or XML-unsafe formula", async () => {
		await reject({ formula: "=A1" }, /stored form/);
		await reject({ formula: "" }, /must not be empty/);
		await reject({ formula: "A".repeat(8193) }, /8192-character limit/);
		await reject({ formula: 123 }, /formula must be a string/);
		await reject({ formula: `A1${String.fromCharCode(1)}` }, /not allowed in XML/);
	});

	it("rejects a cached string result that isn't XML-safe", async () => {
		await reject({ formula: "A1", value: `x${String.fromCharCode(1)}` }, /cached string/);
	});

	it("reads a value-flipping formula getter exactly once (TOCTOU)", async () => {
		let reads = 0;
		const cell = {
			get formula() {
				reads++;
				return reads === 1 ? "A1+1" : `B2${String.fromCharCode(1)}`;
			},
		};
		// The single read is the valid "A1+1"; a second (malicious) read never reaches the output.
		const s = await roundtrip(cell);
		expect(s.formula("A1")).toBe("A1+1");
	});
});

describe("bridge — formula round-trip (e2e)", () => {
	it("carries basic.xlsx's formula as live text", async () => {
		const before = await openXlsx(await loadFixture("basic.xlsx"));
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		const s = after.sheet(after.sheets[0]?.name ?? "");
		expect(s.formula("E1")).toBe("B1*2");
		expect(s.cell("E1").value).toBe(84);
	});

	it("carries error-cell formulas (cached error becomes its string text)", async () => {
		const before = await openXlsx(await loadFixture("errors.xlsx"));
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		const s = after.sheet(after.sheets[0]?.name ?? "");
		expect(s.formula("A1")).toBe("5/0");
		expect(s.cell("A1").value).toBe("#DIV/0!"); // error text preserved as a string result
	});

	it("carries a shared formula as its translated per-cell text", async () => {
		const before = await openXlsx(await loadFixture("shared-formula.xlsx"));
		const after = await openXlsx(await writeXlsx(await workbookToInput(before)));
		const s = after.sheet("Calc");
		expect(s.formula("B2")).toBe("A2*2"); // was a shared dependent; now a plain translated formula
		expect(s.formula("D1")).toBe("A1:A3*2"); // array master text preserved as a plain formula
	});
});
