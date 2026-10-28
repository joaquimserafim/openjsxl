import { describe, expect, it } from "vitest";
import type { Cell } from "../../types";
import { holdsValue } from "../eval";

const empty: Cell = { ref: "A1", type: "empty", value: null };

describe("holdsValue", () => {
	it("keeps a cell with a value of any type", () => {
		const cells: readonly Cell[] = [
			{ ref: "A1", type: "number", value: 0 },
			{ ref: "A1", type: "string", value: "" },
			{ ref: "A1", type: "boolean", value: false },
			{ ref: "A1", type: "date", value: new Date(Date.UTC(2024, 0, 1)) },
			{ ref: "A1", type: "error", value: "#N/A" },
		];
		for (const cell of cells) expect(holdsValue(cell, undefined)).toBe(true);
	});

	it("keeps a formula cell, even one that reads as empty", () => {
		expect(holdsValue(empty, "B1*2")).toBe(true);
		expect(holdsValue(empty, "")).toBe(true);
	});

	it("drops an empty cell without a formula", () => {
		expect(holdsValue(empty, undefined)).toBe(false);
	});
});
