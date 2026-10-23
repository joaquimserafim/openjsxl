import { describe, expect, it } from "vitest";
import {
	passingModule,
	recordingTarget,
	reportedModule,
	reportedTest,
	testSummary,
} from "./test-count-fixtures";

describe("guard reporter test doubles", () => {
	it("builds a reported test with its module, full name and state", () => {
		const test = reportedTest("m.test.ts", "suite > case", "skipped");
		expect(test.fullName).toBe("suite > case");
		expect(test.module.relativeModuleId).toBe("m.test.ts");
		expect(test.result().state).toBe("skipped");
	});

	it("yields a module's tests in order, on every traversal", () => {
		const tests = [
			reportedTest("m.test.ts", "a", "passed"),
			reportedTest("m.test.ts", "b", "failed"),
		];
		const module = reportedModule(tests);
		expect([...module.children.allTests()]).toEqual(tests);
		expect([...module.children.allTests()]).toEqual(tests);
	});

	it("builds a module of passing tests", () => {
		const states = [...passingModule("m.test.ts", 3).children.allTests()].map(
			(test) => test.result().state,
		);
		expect(states).toEqual(["passed", "passed", "passed"]);
	});

	it("records report output and the exit code", () => {
		const output = recordingTarget(2);
		output.target.writeSuccess("ok");
		output.target.writeFailure("bad");
		output.target.exitCode = 1;
		expect(output.successes).toEqual(["ok"]);
		expect(output.failures).toEqual(["bad"]);
		expect(output.target.exitCode).toBe(1);
	});

	it("fills unspecified summary counts with empty values", () => {
		expect(testSummary({ executed: 4 })).toEqual({
			collected: 0,
			executed: 4,
			allowedSkips: 0,
			unexpected: [],
		});
	});
});
