import type {
	MinimumTestCountReportTarget,
	ReportedTest,
	ReportedTestModule,
	TestCountSummary,
} from "../minimum-test-count-reporter";

/** Test doubles for the guard reporter: the structural shapes it reads from Vitest's run. */
export function reportedTest(module: string, fullName: string, state: string): ReportedTest {
	return {
		fullName,
		module: { relativeModuleId: module },
		result: () => ({ state }),
	};
}

export function reportedModule(tests: readonly ReportedTest[]): ReportedTestModule {
	return {
		children: {
			*allTests(): Generator<ReportedTest> {
				yield* tests;
			},
		},
	};
}

export function passingModule(module: string, count: number): ReportedTestModule {
	return reportedModule(
		Array.from({ length: count }, (_, i) => reportedTest(module, `case ${i}`, "passed")),
	);
}

export function recordingTarget(initialExitCode: typeof process.exitCode = undefined): {
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

export function testSummary(partial: Partial<TestCountSummary>): TestCountSummary {
	return { collected: 0, executed: 0, allowedSkips: 0, unexpected: [], ...partial };
}
