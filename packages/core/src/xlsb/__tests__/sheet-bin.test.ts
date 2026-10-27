import { describe, expect, it } from "vitest";
import { parseXlsbSheet } from "../sheet";
import { fullBlank, fullRk, repeatedCellsBin, rowRecord, shortBlank, shortRk } from "./sheet-bin";

const parse = (bytes: readonly number[]) =>
	parseXlsbSheet(Uint8Array.from(bytes), [], undefined, false);

describe("sheet.bin builders", () => {
	it("encode rows, RK numbers and blanks", () => {
		expect(rowRecord(2)).toEqual([0x00, 4, 2, 0, 0, 0]);
		expect(shortRk(7, 2)).toEqual([0x0d, 8, 2, 0, 0, 0, 0x1e, 0, 0, 0]);
		expect(shortBlank()).toEqual([0x0c, 0]);
		expect(shortBlank(3)).toEqual([0x0c, 4, 3, 0, 0, 0]);
		expect(fullRk(1, 7, 2)).toEqual([0x02, 12, 1, 0, 0, 0, 2, 0, 0, 0, 0x1e, 0, 0, 0]);
		expect(fullBlank(1, 3)).toEqual([0x01, 8, 1, 0, 0, 0, 3, 0, 0, 0]);
	});

	it("build records the sheet parser reads back", () => {
		const sheet = parse([
			...rowRecord(1),
			...fullRk(2, 7, 4),
			...shortRk(8),
			...shortBlank(5),
			...fullBlank(6),
		]);
		expect([...sheet.cells.values()]).toEqual([
			{ ref: "C2", type: "number", value: 7 },
			{ ref: "D2", type: "number", value: 8 },
		]);
		expect([...sheet.cellStyles]).toEqual([
			["C2", 4],
			["E2", 5],
		]);
	});

	it("repeat one cell record across rows", () => {
		const bin = repeatedCellsBin(2, 3, shortRk(1));
		expect(bin.length).toBe(2 * (6 + 3 * 10));
		const sheet = parseXlsbSheet(bin, [], undefined, false);
		expect([sheet.cells.size, [...sheet.cells.keys()].at(-1)]).toEqual([6, "C2"]);
		expect(repeatedCellsBin(2, 3, shortBlank()).length).toBe(24);
	});
});
