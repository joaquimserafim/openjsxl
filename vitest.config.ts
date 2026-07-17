import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const resolvePath = (p: string) => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
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
		include: ["packages/*/src/**/*.test.ts", "packages/*/test/**/*.test.ts"],
		// No feature tests land until F0.3/F0.4; keep the skeleton's run green.
		passWithNoTests: true,
	},
});
