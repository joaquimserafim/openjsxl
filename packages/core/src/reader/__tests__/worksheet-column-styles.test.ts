import { describe, expect, it } from "vitest";
import { MAX_COL } from "../../ooxml/a1";
import {
	columnStyleFromToken,
	createColumnStyleResolver,
	nextUnassigned,
} from "../worksheet-column-styles";

describe("worksheet column-style resolver", () => {
	it("finds and path-compresses the next unassigned column", () => {
		const links = Uint32Array.from([0, 2, 3, 3]);
		expect(nextUnassigned(links, 1)).toBe(3);
		expect([...links]).toEqual([0, 3, 3, 3]);
	});

	it("reads only complete integer declarations", () => {
		expect(columnStyleFromToken({ min: "2", max: "4", style: "7" })).toEqual({
			min: 2,
			max: 4,
			style: 7,
		});
		expect(columnStyleFromToken({ min: "2", max: "4" })).toBeUndefined();
		expect(columnStyleFromToken({ min: "x", max: "4", style: "7" })).toBeUndefined();
		expect(columnStyleFromToken({ min: "2.5", max: "4", style: "7" })).toBeUndefined();
	});

	it("keeps the first style for overlaps, including style zero", () => {
		const styles = createColumnStyleResolver();
		expect(styles.add({ min: 1, max: 2, style: 0 })).toBe(2);
		expect(styles.add({ min: 2, max: 3, style: 8 })).toBe(1);

		// Input ranges 1..2/style 0 then 2..3/style 8; output columns 1..3 are 0, 0, 8.
		expect([styles.styleFor(1), styles.styleFor(2), styles.styleFor(3)]).toEqual([0, 0, 8]);
	});

	it("clips partial ranges to the grid and ignores invalid or wholly outside ranges", () => {
		const styles = createColumnStyleResolver();
		expect(styles.add({ min: -100, max: 2, style: 7 })).toBe(2);
		expect(styles.add({ min: MAX_COL, max: Number.MAX_VALUE, style: 8 })).toBe(1);
		expect(styles.add({ min: 5, max: 4, style: 9 })).toBe(0);
		expect(styles.add({ min: MAX_COL + 1, max: MAX_COL + 2, style: 9 })).toBe(0);
		expect(styles.add({ min: 3.5, max: 4, style: 9 })).toBe(0);
		expect(styles.add({ min: 3, max: 3, style: 4_294_967_295 })).toBe(1);
		expect(styles.add({ min: 4, max: 4, style: 1.5 })).toBe(0);
		expect(styles.add({ min: 4, max: 4, style: -7 })).toBe(1);

		expect(styles.styleFor(1)).toBe(7);
		expect(styles.styleFor(2)).toBe(7);
		expect(styles.styleFor(3)).toBe(4_294_967_295); // JS integer values are not Int32-truncated
		expect(styles.styleFor(4)).toBe(-7); // a rejected declaration did not reserve the slot
		expect(styles.styleFor(MAX_COL)).toBe(8);
		expect(styles.styleFor(0)).toBeUndefined();
		expect(styles.styleFor(MAX_COL + 1)).toBeUndefined();
	});

	it("assigns at most the fixed grid and never reads declarations during lookup", () => {
		let propertyReads = 0;
		const fullGrid = {
			get min() {
				propertyReads++;
				return 1;
			},
			get max() {
				propertyReads++;
				return MAX_COL;
			},
			get style() {
				propertyReads++;
				return 7;
			},
		};
		const styles = createColumnStyleResolver();
		expect(styles.add(fullGrid)).toBe(MAX_COL);
		expect(propertyReads).toBe(3);

		// These attacker-controlled overlaps cannot assign more work or replace the first style.
		let newlyAssigned = 0;
		for (let i = 0; i < 20_000; i++) {
			newlyAssigned += styles.add({ min: 1, max: MAX_COL, style: 8 });
		}
		expect(newlyAssigned).toBe(0);

		let mismatches = 0;
		for (let column = 1; column <= MAX_COL; column++) {
			if (styles.styleFor(column) !== 7) mismatches++;
		}
		expect(mismatches).toBe(0);
		expect(propertyReads).toBe(3);
	});
});
