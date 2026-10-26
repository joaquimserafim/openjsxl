import { describe, expect, it } from "vitest";
import { MAX_CSV_CELLS } from "../../csv";
import { MAX_COL, MAX_ROW } from "../../ooxml";
import { csvCells, openCsv } from "../csv";

// F11.58: openCsv decodes through decodeText and places fields straight from the scanner, against a
// cell budget below V8's Map limit. Every overflow fails with XlsxError('part-too-large').

const typedTooLarge = expect.objectContaining({ name: "XlsxError", code: "part-too-large" });
const budgetError = expect.objectContaining({
	name: "XlsxError",
	code: "part-too-large",
	message: expect.stringContaining("cell budget"),
});

describe("csvCells", () => {
	it("places typed cells, keyed by ref, with their dimension", () => {
		const { cells, dimension } = csvCells("a,1\n,TRUE", ",", true);
		expect([...cells.values()]).toEqual([
			{ ref: "A1", type: "string", value: "a" },
			{ ref: "B1", type: "number", value: 1 },
			{ ref: "B2", type: "boolean", value: true },
		]);
		expect(dimension).toBe("A1:B2");
	});

	it("fails typed once the non-empty cells pass the budget", () => {
		expect(() => csvCells("a,b,c", ",", true, 2)).toThrow(budgetError);
		expect(() => csvCells("a,b,c", ",", false, 2)).toThrow(budgetError);
	});

	it("can fill the budget exactly", () => {
		expect(csvCells("a,b,c", ",", true, 3).cells.size).toBe(3);
	});

	it("charges nothing for empty fields", () => {
		const { cells, dimension } = csvCells(`${",".repeat(5_000)}a\n\n\r\n,,`, ",", true, 1);
		expect([cells.size, dimension]).toEqual([1, "GJI1:GJI1"]);
	});

	it("charges nothing for fields past the last column or row, which are dropped", () => {
		const wide = csvCells(
			Array(MAX_COL + 5)
				.fill("1")
				.join(","),
			",",
			true,
			MAX_COL,
		);
		expect([wide.cells.size, wide.dimension]).toEqual([MAX_COL, "A1:XFD1"]);
		const tall = csvCells(`${"\n".repeat(MAX_ROW)}a`, ",", true, 0);
		expect([tall.cells.size, tall.dimension]).toEqual([0, undefined]);
	});
});

describe("openCsv limits", () => {
	it("keeps the cell budget below V8's 2^24-entry Map limit", () => {
		expect(MAX_CSV_CELLS).toBeLessThan(2 ** 24);
	});

	it("fails typed on bytes whose text passes the maximum string length", () => {
		// V8 caps a string at 0x1fffffe8 chars. The decode fails before it returns: the test holds the
		// 512 MiB input array, plus the decoder's buffer while it runs.
		const bytes = new Uint8Array(0x1fffffe8 + 8).fill(0x41);
		expect(() => openCsv(bytes)).toThrow(typedTooLarge);
	});

	it("reads ArrayBuffer input as its bytes", () => {
		const bytes = new TextEncoder().encode("\ufeffa;1");
		const s = openCsv(bytes.buffer).sheet("Sheet1");
		expect([s.cell("A1").value, s.cell("B1").value]).toEqual(["a", 1]);
	});
});
