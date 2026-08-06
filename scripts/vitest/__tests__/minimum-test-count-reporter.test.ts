import { describe, expect, it, vi } from "vitest";
import MinimumTestCountReporter, {
	countCollectedTests,
	MINIMUM_COLLECTED_TESTS,
	type MinimumTestCountReportTarget,
	minimumTestCountFailure,
	reportMinimumTestCount,
} from "../minimum-test-count-reporter";

function collectedModule(testCount: number): {
	readonly children: { allTests(): Iterable<unknown> };
} {
	return {
		children: {
			*allTests(): Generator<unknown> {
				for (let i = 0; i < testCount; i++) yield { id: i };
			},
		},
	};
}

function recordingTarget(initialExitCode: typeof process.exitCode = undefined): {
	readonly target: MinimumTestCountReportTarget;
	readonly successes: string[];
	readonly failures: string[];
} {
	const successes: string[] = [];
	const failures: string[] = [];
	const target: MinimumTestCountReportTarget = {
		exitCode: initialExitCode,
		writeSuccess(message: string): void {
			successes.push(message);
		},
		writeFailure(message: string): void {
			failures.push(message);
		},
	};
	return { target, successes, failures };
}

describe("minimum test-count reporter", () => {
	it("counts tests across modules, including every yielded test state", () => {
		expect(countCollectedTests([collectedModule(2), collectedModule(3)])).toBe(5);
	});

	it("accepts the configured floor and anything above it", () => {
		expect(minimumTestCountFailure(MINIMUM_COLLECTED_TESTS)).toBeUndefined();
		expect(minimumTestCountFailure(MINIMUM_COLLECTED_TESTS + 1)).toBeUndefined();
	});

	it("returns a clear failure below the floor", () => {
		expect(minimumTestCountFailure(899)).toBe(
			"Test discovery guard: collected 899 tests; expected at least 900.",
		);
	});

	it("rejects invalid counts and a disabled floor", () => {
		expect(() => minimumTestCountFailure(-1)).toThrow(TypeError);
		expect(() => minimumTestCountFailure(1.5)).toThrow(TypeError);
		expect(() => minimumTestCountFailure(0, 0)).toThrow(TypeError);
	});

	it("sets a successful process to failure when the floor is missed", () => {
		const output = recordingTarget();
		reportMinimumTestCount(899, 900, output.target);

		expect(output.target.exitCode).toBe(1);
		expect(output.failures).toEqual([
			"\nTest discovery guard: collected 899 tests; expected at least 900.\n",
		]);
		expect(output.successes).toEqual([]);
	});

	it("does not replace an existing nonzero exit code", () => {
		const output = recordingTarget(2);
		reportMinimumTestCount(899, 900, output.target);

		expect(output.target.exitCode).toBe(2);
		expect(output.failures).toHaveLength(1);
	});

	it("reports success without changing the exit code", () => {
		const output = recordingTarget();
		reportMinimumTestCount(900, 900, output.target);

		expect(output.target.exitCode).toBeUndefined();
		expect(output.successes).toEqual([
			"\nTest discovery guard: 900 collected (minimum 900).\n",
		]);
		expect(output.failures).toEqual([]);
	});

	it("wires lifecycle failures to the process status and stderr", () => {
		const previousExitCode = process.exitCode;
		const writeFailure = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
		try {
			process.exitCode = undefined;
			new MinimumTestCountReporter({ minimum: 1 }).onTestRunEnd([], [], "passed");

			expect(process.exitCode).toBe(1);
			expect(writeFailure).toHaveBeenCalledWith(
				"\nTest discovery guard: collected 0 tests; expected at least 1.\n",
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
