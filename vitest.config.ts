import { fileURLToPath } from "node:url";
import { defineConfig, type Plugin } from "vitest/config";
import MinimumTestCountReporter from "./scripts/vitest/minimum-test-count-reporter";

const resolvePath = (p: string) => fileURLToPath(new URL(p, import.meta.url));
const minimumTestCountPlugin: Plugin = {
	name: "openjsxl-minimum-test-count",
	configureVitest({ vitest }): void {
		if (process.env.npm_lifecycle_event !== "test") return;
		// This hook runs after CLI options are resolved, so `--reporter` cannot replace the guard.
		// Vitest adds CI annotations automatically unless a CLI reporter replaced its defaults.
		const hasGitHubReporter = vitest.config.reporters.some(
			(reporter) => Array.isArray(reporter) && reporter[0] === "github-actions",
		);
		if (process.env.GITHUB_ACTIONS === "true" && !hasGitHubReporter) {
			vitest.config.reporters.push(["github-actions", {}]);
		}
		vitest.config.reporters.push(new MinimumTestCountReporter());
	},
};

export default defineConfig({
	plugins: [minimumTestCountPlugin],
	resolve: {
		// Run tests against TypeScript source — no build step required in dev. Subpath aliases must
		// precede their bare parents: an alias key matches a specifier that equals it OR starts with
		// `key + "/"`, so a bare `openjsxl` would otherwise swallow `openjsxl/formula`.
		alias: {
			"@openjsxl/core/formula": resolvePath("./packages/core/src/formula/index.ts"),
			"@openjsxl/core": resolvePath("./packages/core/src/index.ts"),
			"openjsxl/formula": resolvePath("./packages/openjsxl/src/formula.ts"),
			openjsxl: resolvePath("./packages/openjsxl/src/index.ts"),
		},
	},
	test: {
		include: ["packages/*/src/**/*.test.ts", "scripts/**/*.test.ts"],
		// `pnpm test` is the full local/CI gate. Focused direct Vitest and the two-test long fuzz run
		// deliberately omit the executed-test floor and skip allowlist, while no-test discovery still
		// fails for every command.
	},
});
