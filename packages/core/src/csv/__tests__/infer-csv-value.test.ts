import { afterEach, describe, expect, it, vi } from "vitest";
import { inferCsvValue } from "../parse";

// F11.58: only a 4- or 5-unit field can be a boolean, so only those are lowercased. Lowercasing a huge
// field of U+0130 (two units in lowercase) would pass the maximum string length and kill the process.

afterEach(() => {
	vi.restoreAllMocks();
});

describe("inferCsvValue's boolean check", () => {
	it("lowercases only fields of 4 or 5 units", () => {
		const lower = vi.spyOn(String.prototype, "toLowerCase");
		inferCsvValue("x".repeat(100_000));
		inferCsvValue("abc");
		inferCsvValue("abcdef");
		expect(lower).not.toHaveBeenCalled();
		inferCsvValue("abcd");
		inferCsvValue("abcde");
		expect(lower).toHaveBeenCalledTimes(2);
	});

	it("reads TRUE and FALSE in any case, and nothing longer or shorter", () => {
		expect(inferCsvValue("tRuE")).toEqual({ type: "boolean", value: true });
		expect(inferCsvValue("fALSe")).toEqual({ type: "boolean", value: false });
		expect(inferCsvValue(" true")).toEqual({ type: "string", value: " true" });
		expect(inferCsvValue("TRUEx")).toEqual({ type: "string", value: "TRUEx" });
		expect(inferCsvValue("İ".repeat(3))).toEqual({ type: "string", value: "İİİ" });
	});
});
