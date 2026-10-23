import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import { isDateFormatCode } from "../../ooxml/styles";
import { parseXlsbStyles } from "../parts";
import { stylesBin } from "./styles-bin";

// F11.50: xlsb shares the xlsx date decision (linear, cached per id) and the 255-character
// number-format limit: a longer BrtFmt code reads as absent and never as a date.
const DATE_255 = `yyyy${"0".repeat(251)}`;
const DATE_256 = `${DATE_255}0`;

describe("xlsb number-format limit and date decision", () => {
	it("keeps a 255-character code and types it as a date", () => {
		const styles = parseXlsbStyles(stylesBin([[164, DATE_255]], [0, 164]));
		expect(styles.formatCode(1)).toBe(DATE_255);
		expect(styles.isDateStyle(1)).toBe(true);
	});

	it("drops a 256-character code to absent", () => {
		const styles = parseXlsbStyles(stylesBin([[164, DATE_256]], [0, 164]));
		expect(styles.formatCode(1)).toBeUndefined();
		expect(styles.isDateStyle(1)).toBe(false);
	});

	it("drops an oversized redefinition of a built-in id without falling back to the built-in", () => {
		const styles = parseXlsbStyles(stylesBin([[14, DATE_256]], [0, 14]));
		expect(styles.formatCode(1)).toBeUndefined();
		expect(styles.isDateStyle(1)).toBe(false);
	});

	it("keeps built-in ids and repeated lookups stable", () => {
		const styles = parseXlsbStyles(stylesBin([], [0, 14, 2]));
		expect([styles.isDateStyle(1), styles.isDateStyle(1), styles.isDateStyle(2)]).toEqual([
			true,
			true,
			false,
		]);
		expect(styles.formatCode(2)).toBe("0.00");
	});

	it("decides each custom format id once however many cells and xfs use it", () => {
		let calls = 0;
		const styles = parseXlsbStyles(
			stylesBin(
				[
					[164, "yyyy-mm-dd"],
					[165, "0.0"],
				],
				[0, 164, 165, 164, 165],
			),
			(code) => {
				calls++;
				return isDateFormatCode(code);
			},
		);
		const decisions: boolean[] = [];
		for (let cell = 0; cell < 100; cell++) decisions.push(styles.isDateStyle(1 + (cell % 4)));
		expect(decisions.slice(0, 4)).toEqual([true, false, true, false]);
		expect(calls).toBe(2);
	});

	it("is linear in code length on unclosed brackets", () => {
		const ratio = bestDoublingRatio((n) => {
			const bytes = stylesBin([[164, "[".repeat(n)]], [0, 164]);
			return () => parseXlsbStyles(bytes).isDateStyle(1);
		}, 25_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 60_000);
});
