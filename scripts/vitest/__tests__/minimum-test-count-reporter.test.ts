import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import MinimumTestCountReporter, {
	ALLOWED_SKIPS,
	type AllowedSkip,
	isAllowedSkip,
	MINIMUM_EXECUTED_TESTS,
	reportTestCount,
	requireMinimum,
	requireNonNegativeInteger,
	summarizeTests,
	testCountFailure,
} from "../minimum-test-count-reporter";
import {
	passingModule,
	recordingTarget,
	reportedModule,
	reportedTest,
	testSummary,
} from "./test-count-fixtures";

const ROOT = new URL("../../../", import.meta.url);

const SUITE_ALLOW: readonly AllowedSkip[] = [{ module: "a.test.ts", suite: "real (local)" }];

describe("requireNonNegativeInteger", () => {
	it("accepts zero and positive safe integers", () => {
		expect(() => requireNonNegativeInteger(0, "n")).not.toThrow();
		expect(() => requireNonNegativeInteger(2071, "n")).not.toThrow();
	});

	it("rejects negative, fractional and unsafe values", () => {
		expect(() => requireNonNegativeInteger(-1, "n")).toThrow(TypeError);
		expect(() => requireNonNegativeInteger(1.5, "n")).toThrow(TypeError);
		expect(() => requireNonNegativeInteger(Number.MAX_SAFE_INTEGER + 1, "n")).toThrow(
			TypeError,
		);
	});
});

describe("requireMinimum", () => {
	it("rejects a disabled or invalid floor", () => {
		expect(() => requireMinimum(0)).toThrow("minimum test count must be greater than zero");
		expect(() => requireMinimum(-5)).toThrow(TypeError);
		expect(() => requireMinimum(1)).not.toThrow();
	});
});

describe("isAllowedSkip", () => {
	it("allows only the named suite and its descendants within the module", () => {
		const inSuite = reportedTest("a.test.ts", "real (local) > reads notes", "skipped");
		const suiteItself = reportedTest("a.test.ts", "real (local)", "skipped");
		const sibling = reportedTest("a.test.ts", "parses notes", "skipped");
		const prefixOnly = reportedTest("a.test.ts", "real (local)x > case", "skipped");
		const otherModule = reportedTest("b.test.ts", "real (local) > reads notes", "skipped");

		expect(isAllowedSkip(inSuite, SUITE_ALLOW)).toBe(true);
		expect(isAllowedSkip(suiteItself, SUITE_ALLOW)).toBe(true);
		expect(isAllowedSkip(sibling, SUITE_ALLOW)).toBe(false);
		expect(isAllowedSkip(prefixOnly, SUITE_ALLOW)).toBe(false);
		expect(isAllowedSkip(otherModule, SUITE_ALLOW)).toBe(false);
	});

	it("names the two environment-gated suites that exist in the repository", () => {
		expect(ALLOWED_SKIPS).toEqual([
			{ module: "packages/fuzz/src/__tests__/long-run.test.ts", suite: "F9.4 long run" },
			{
				module: "packages/core/src/reader/__tests__/comments.test.ts",
				suite: "Worksheet.comments — real Excel (local)",
			},
		]);
		for (const entry of ALLOWED_SKIPS) {
			const path = new URL(entry.module, ROOT);
			expect(existsSync(path)).toBe(true);
			expect(readFileSync(path, "utf8")).toContain(`"${entry.suite}"`);
		}
	});

	it("does not allow a stray skip outside the gated suite of an allowed module", () => {
		const stray = reportedTest(
			"packages/fuzz/src/__tests__/long-run.test.ts",
			"unrelated > case",
			"skipped",
		);
		expect(isAllowedSkip(stray)).toBe(false);
	});
});

describe("summarizeTests", () => {
	it("counts passed and failed tests as executed across modules", () => {
		const result = summarizeTests(
			[
				passingModule("x.test.ts", 2),
				reportedModule([reportedTest("y.test.ts", "breaks", "failed")]),
			],
			SUITE_ALLOW,
		);
		expect(result).toEqual({ collected: 3, executed: 3, allowedSkips: 0, unexpected: [] });
	});

	it("separates allowed skips from unexpected skipped, todo and pending tests", () => {
		const result = summarizeTests(
			[
				reportedModule([
					reportedTest("a.test.ts", "real (local) > reads notes", "skipped"),
					reportedTest("a.test.ts", "parses notes", "skipped"),
					reportedTest("a.test.ts", "never finished", "pending"),
					reportedTest("a.test.ts", "runs", "passed"),
				]),
			],
			SUITE_ALLOW,
		);
		expect(result).toEqual({
			collected: 4,
			executed: 1,
			allowedSkips: 1,
			unexpected: ["a.test.ts > parses notes", "a.test.ts > never finished"],
		});
	});

	it("treats a filter that skips everything as unexpected skips, not executed tests", () => {
		const everythingSkipped = reportedModule(
			Array.from({ length: 3 }, (_, i) => reportedTest("z.test.ts", `case ${i}`, "skipped")),
		);
		const result = summarizeTests([everythingSkipped], SUITE_ALLOW);
		expect(result.executed).toBe(0);
		expect(result.unexpected).toHaveLength(3);
	});
});

describe("testCountFailure", () => {
	it("accepts the floor and anything above it with no unexpected skips", () => {
		expect(testCountFailure(testSummary({ collected: 2071, executed: 2000 }))).toBeUndefined();
		expect(testCountFailure(testSummary({ collected: 2071, executed: 2069 }))).toBeUndefined();
	});

	it("reports an executed count below the floor with the collected count", () => {
		expect(testCountFailure(testSummary({ collected: 2071, executed: 1311 }), 2000)).toBe(
			"Test discovery guard: executed 1311 tests; expected at least 2000 (2071 collected).",
		);
	});

	it("lists the first unexpected skips and counts the rest", () => {
		const unexpected = Array.from({ length: 7 }, (_, i) => `m.test.ts > case ${i}`);
		expect(testCountFailure(testSummary({ collected: 9, executed: 2, unexpected }), 2)).toBe(
			"Test discovery guard: 7 tests skipped or pending outside the allowlist: m.test.ts > case 0, m.test.ts > case 1, m.test.ts > case 2, m.test.ts > case 3, m.test.ts > case 4, and 2 more.",
		);
	});

	it("uses the singular for a count of one", () => {
		expect(testCountFailure(testSummary({ collected: 1, executed: 1 }), 2)).toBe(
			"Test discovery guard: executed 1 test; expected at least 2 (1 collected).",
		);
	});

	it("reports both problems together", () => {
		expect(
			testCountFailure(
				testSummary({ collected: 3, executed: 0, unexpected: ["m.test.ts > a"] }),
				1,
			),
		).toBe(
			"Test discovery guard: executed 0 tests; expected at least 1 (3 collected); 1 test skipped or pending outside the allowlist: m.test.ts > a.",
		);
	});

	it("rejects invalid counts and a disabled floor", () => {
		expect(() => testCountFailure(testSummary({ executed: -1 }))).toThrow(TypeError);
		expect(() => testCountFailure(testSummary({ collected: 1.5 }))).toThrow(TypeError);
		expect(() => testCountFailure(testSummary({ executed: 1 }), 0)).toThrow(TypeError);
	});

	it("defaults to the full-suite floor", () => {
		expect(MINIMUM_EXECUTED_TESTS).toBe(2000);
		expect(testCountFailure(testSummary({ executed: MINIMUM_EXECUTED_TESTS - 1 }))).toContain(
			"expected at least 2000",
		);
	});
});

describe("reportTestCount", () => {
	it("sets a successful process to failure when the guard is not met", () => {
		const output = recordingTarget();
		reportTestCount(testSummary({ collected: 900, executed: 899 }), 900, output.target);

		expect(output.target.exitCode).toBe(1);
		expect(output.failures).toEqual([
			"\nTest discovery guard: executed 899 tests; expected at least 900 (900 collected).\n",
		]);
		expect(output.successes).toEqual([]);
	});

	it("does not replace an existing nonzero exit code", () => {
		const output = recordingTarget(2);
		reportTestCount(testSummary({ collected: 1, executed: 0 }), 1, output.target);

		expect(output.target.exitCode).toBe(2);
		expect(output.failures).toHaveLength(1);
	});

	it("reports executed, allowed-skip and collected counts on success", () => {
		const output = recordingTarget();
		reportTestCount(
			testSummary({ collected: 2071, executed: 2068, allowedSkips: 3 }),
			2000,
			output.target,
		);

		expect(output.target.exitCode).toBeUndefined();
		expect(output.successes).toEqual([
			"\nTest discovery guard: 2068 executed, 3 allowed skips, 2071 collected (minimum 2000).\n",
		]);
		expect(output.failures).toEqual([]);
	});
});

describe("MinimumTestCountReporter", () => {
	it("rejects an invalid configured floor", () => {
		expect(() => new MinimumTestCountReporter({ minimum: 0 })).toThrow(TypeError);
	});

	it("wires lifecycle failures to the process status and stderr", () => {
		const previousExitCode = process.exitCode;
		const writeFailure = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
		try {
			process.exitCode = undefined;
			new MinimumTestCountReporter({ minimum: 1 }).onTestRunEnd([], [], "passed");

			expect(process.exitCode).toBe(1);
			expect(writeFailure).toHaveBeenCalledWith(
				"\nTest discovery guard: executed 0 tests; expected at least 1 (0 collected).\n",
			);
		} finally {
			process.exitCode = previousExitCode;
			writeFailure.mockRestore();
		}
	});

	it("fails a run whose filter skipped every collected test", () => {
		const previousExitCode = process.exitCode;
		const writeFailure = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
		try {
			process.exitCode = undefined;
			const skipped = reportedModule([reportedTest("x.test.ts", "case", "skipped")]);
			new MinimumTestCountReporter({ minimum: 1, allowedSkips: [] }).onTestRunEnd(
				[skipped],
				[],
				"passed",
			);

			expect(process.exitCode).toBe(1);
			expect(writeFailure).toHaveBeenCalledWith(
				"\nTest discovery guard: executed 0 tests; expected at least 1 (1 collected); 1 test skipped or pending outside the allowlist: x.test.ts > case.\n",
			);
		} finally {
			process.exitCode = previousExitCode;
			writeFailure.mockRestore();
		}
	});

	it("leaves an interrupted lifecycle run untouched", () => {
		const previousExitCode = process.exitCode;
		const writeFailure = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
		try {
			process.exitCode = 2;
			new MinimumTestCountReporter({ minimum: 1 }).onTestRunEnd([], [], "interrupted");

			expect(process.exitCode).toBe(2);
			expect(writeFailure).not.toHaveBeenCalled();
		} finally {
			process.exitCode = previousExitCode;
			writeFailure.mockRestore();
		}
	});
});
