import type { Reporter } from "vitest/reporters";

/**
 * Full-suite floor. On 2026-10-23 the gate executed 2,087 tests in a clone without local fixtures
 * (as in CI) and 2,088 with the local comments fixture.
 */
export const MINIMUM_EXECUTED_TESTS = 2000;

export interface AllowedSkip {
	/** Test module path relative to the project root. */
	readonly module: string;
	/** Top-level suite whose tests may skip. */
	readonly suite: string;
}

/**
 * The only tests the full gate may skip. Both are environment-gated: the long fuzz run needs
 * `FUZZ_LONG`, and the real-Excel comments suite needs a local-only fixture absent from CI.
 */
export const ALLOWED_SKIPS: readonly AllowedSkip[] = [
	{ module: "packages/fuzz/src/__tests__/long-run.test.ts", suite: "F9.4 long run" },
	{
		module: "packages/core/src/reader/__tests__/comments.test.ts",
		suite: "Worksheet.comments — real Excel (local)",
	},
];

export interface ReportedTest {
	readonly fullName: string;
	readonly module: { readonly relativeModuleId: string };
	result(): { readonly state: string };
}

export interface ReportedTestModule {
	readonly children: {
		allTests(): Iterable<ReportedTest>;
	};
}

export interface TestCountSummary {
	readonly collected: number;
	readonly executed: number;
	readonly allowedSkips: number;
	/** `module > full test name` for every test that neither ran nor matched an allowed skip. */
	readonly unexpected: readonly string[];
}

interface MinimumTestCountReporterOptions {
	readonly minimum?: number;
	readonly allowedSkips?: readonly AllowedSkip[];
}

export interface MinimumTestCountReportTarget {
	exitCode: typeof process.exitCode;
	writeSuccess(message: string): void;
	writeFailure(message: string): void;
}

type TestRunEndArgs = Parameters<NonNullable<Reporter["onTestRunEnd"]>>;

/** Names listed in a failure before the remainder is summarized as a count. */
const UNEXPECTED_PREVIEW = 5;

export function requireNonNegativeInteger(value: number, name: string): void {
	if (!Number.isSafeInteger(value) || value < 0) {
		throw new TypeError(`${name} must be a non-negative safe integer`);
	}
}

export function requireMinimum(minimum: number): void {
	requireNonNegativeInteger(minimum, "minimum test count");
	if (minimum === 0) throw new TypeError("minimum test count must be greater than zero");
}

/** Whether a skipped test belongs to one of the allowed environment-gated suites. */
export function isAllowedSkip(
	test: ReportedTest,
	allowed: readonly AllowedSkip[] = ALLOWED_SKIPS,
): boolean {
	return allowed.some((entry) => {
		if (test.module.relativeModuleId !== entry.module) return false;
		return test.fullName === entry.suite || test.fullName.startsWith(`${entry.suite} > `);
	});
}

/** Count executed (passed or failed) tests and every skip outside the allowlist, in one pass. */
export function summarizeTests(
	testModules: readonly ReportedTestModule[],
	allowed: readonly AllowedSkip[] = ALLOWED_SKIPS,
): TestCountSummary {
	let collected = 0;
	let executed = 0;
	let allowedSkips = 0;
	const unexpected: string[] = [];
	for (const testModule of testModules) {
		for (const test of testModule.children.allTests()) {
			collected++;
			const state = test.result().state;
			if (state === "passed" || state === "failed") {
				executed++;
			} else if (state === "skipped" && isAllowedSkip(test, allowed)) {
				allowedSkips++;
			} else {
				unexpected.push(`${test.module.relativeModuleId} > ${test.fullName}`);
			}
		}
	}
	return { collected, executed, allowedSkips, unexpected };
}

/** Return the CI-facing failure text, or `undefined` when the run meets the floor cleanly. */
export function testCountFailure(
	summary: TestCountSummary,
	minimum = MINIMUM_EXECUTED_TESTS,
): string | undefined {
	requireNonNegativeInteger(summary.executed, "executed test count");
	requireNonNegativeInteger(summary.collected, "collected test count");
	requireMinimum(minimum);
	const problems: string[] = [];
	if (summary.executed < minimum) {
		problems.push(
			`executed ${summary.executed} ${summary.executed === 1 ? "test" : "tests"}; expected at least ${minimum} (${summary.collected} collected)`,
		);
	}
	const count = summary.unexpected.length;
	if (count > 0) {
		const shown = summary.unexpected.slice(0, UNEXPECTED_PREVIEW).join(", ");
		const rest = count > UNEXPECTED_PREVIEW ? `, and ${count - UNEXPECTED_PREVIEW} more` : "";
		const noun = count === 1 ? "test" : "tests";
		problems.push(`${count} ${noun} skipped or pending outside the allowlist: ${shown}${rest}`);
	}
	if (problems.length === 0) return undefined;
	return `Test discovery guard: ${problems.join("; ")}.`;
}

/** Report the result and fail a previously successful process when the guard is not met. */
export function reportTestCount(
	summary: TestCountSummary,
	minimum: number,
	target: MinimumTestCountReportTarget,
): void {
	const failure = testCountFailure(summary, minimum);
	if (failure === undefined) {
		target.writeSuccess(
			`\nTest discovery guard: ${summary.executed} executed, ${summary.allowedSkips} allowed skips, ${summary.collected} collected (minimum ${minimum}).\n`,
		);
		return;
	}

	target.writeFailure(`\n${failure}\n`);
	if (
		target.exitCode === undefined ||
		target.exitCode === null ||
		target.exitCode === 0 ||
		target.exitCode === "0"
	) {
		target.exitCode = 1;
	}
}

const processReportTarget: MinimumTestCountReportTarget = {
	get exitCode(): typeof process.exitCode {
		return process.exitCode;
	},
	set exitCode(value: typeof process.exitCode) {
		process.exitCode = value;
	},
	writeSuccess(message: string): void {
		process.stdout.write(message);
	},
	writeFailure(message: string): void {
		process.stderr.write(message);
	},
};

/**
 * Vitest reporter used by the root `pnpm test` gate. It observes the completed run, so the guard
 * adds no second collection or execution pass. Focused/direct Vitest commands intentionally omit it.
 */
export default class MinimumTestCountReporter implements Reporter {
	readonly #minimum: number;
	readonly #allowedSkips: readonly AllowedSkip[];

	constructor(options: MinimumTestCountReporterOptions = {}) {
		this.#minimum = options.minimum ?? MINIMUM_EXECUTED_TESTS;
		this.#allowedSkips = options.allowedSkips ?? ALLOWED_SKIPS;
		requireMinimum(this.#minimum);
	}

	onTestRunEnd(
		testModules: readonly ReportedTestModule[],
		_unhandledErrors: TestRunEndArgs[1],
		reason: TestRunEndArgs[2],
	): void {
		// Preserve Vitest's interruption status rather than replacing it with the discovery failure.
		if (reason === "interrupted") return;
		const summary = summarizeTests(testModules, this.#allowedSkips);
		reportTestCount(summary, this.#minimum, processReportTarget);
	}
}
