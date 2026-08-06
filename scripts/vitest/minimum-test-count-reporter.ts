import type { Reporter } from "vitest/reporters";

/** Current suite: 1,049 passing + 2 skipped before this guard's own tests. */
export const MINIMUM_COLLECTED_TESTS = 900;

interface CollectedTestModule {
	readonly children: {
		allTests(): Iterable<unknown>;
	};
}

interface MinimumTestCountReporterOptions {
	readonly minimum?: number;
}

export interface MinimumTestCountReportTarget {
	exitCode: typeof process.exitCode;
	writeSuccess(message: string): void;
	writeFailure(message: string): void;
}

type TestRunEndArgs = Parameters<NonNullable<Reporter["onTestRunEnd"]>>;

function requireNonNegativeInteger(value: number, name: string): void {
	if (!Number.isSafeInteger(value) || value < 0) {
		throw new TypeError(`${name} must be a non-negative safe integer`);
	}
}

/** Count every collected test, including skipped/todo cases, without running the suite again. */
export function countCollectedTests(testModules: readonly CollectedTestModule[]): number {
	let count = 0;
	for (const testModule of testModules) {
		for (const _test of testModule.children.allTests()) count++;
	}
	return count;
}

/** Return the CI-facing failure text, or `undefined` when discovery meets the floor. */
export function minimumTestCountFailure(
	collected: number,
	minimum = MINIMUM_COLLECTED_TESTS,
): string | undefined {
	requireNonNegativeInteger(collected, "collected test count");
	requireNonNegativeInteger(minimum, "minimum test count");
	if (minimum === 0) throw new TypeError("minimum test count must be greater than zero");
	if (collected >= minimum) return undefined;
	return `Test discovery guard: collected ${collected} tests; expected at least ${minimum}.`;
}

/** Report the result and fail a previously successful process when the floor is missed. */
export function reportMinimumTestCount(
	collected: number,
	minimum: number,
	target: MinimumTestCountReportTarget,
): void {
	const failure = minimumTestCountFailure(collected, minimum);
	if (failure === undefined) {
		target.writeSuccess(
			`\nTest discovery guard: ${collected} collected (minimum ${minimum}).\n`,
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
 * Vitest reporter used by the root `pnpm test` gate. It observes the completed run, so the floor
 * adds no second collection or execution pass. Focused/direct Vitest commands intentionally omit it.
 */
export default class MinimumTestCountReporter implements Reporter {
	readonly #minimum: number;

	constructor(options: MinimumTestCountReporterOptions = {}) {
		this.#minimum = options.minimum ?? MINIMUM_COLLECTED_TESTS;
		minimumTestCountFailure(this.#minimum, this.#minimum);
	}

	onTestRunEnd(
		testModules: TestRunEndArgs[0],
		_unhandledErrors: TestRunEndArgs[1],
		reason: TestRunEndArgs[2],
	): void {
		// Preserve Vitest's interruption status rather than replacing it with the discovery failure.
		if (reason === "interrupted") return;
		const collected = countCollectedTests(testModules);
		reportMinimumTestCount(collected, this.#minimum, processReportTarget);
	}
}
