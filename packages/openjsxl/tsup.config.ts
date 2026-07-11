import { defineConfig } from "tsup";

export default defineConfig({
	// Mirror core's two entries: the facade "." and the opt-in "./formula" (M8).
	entry: { index: "src/index.ts", formula: "src/formula.ts" },
	format: ["esm"],
	target: "node24",
	dts: true,
	clean: true,
	treeshake: true,
	// One self-contained file per entry (see @openjsxl/core's tsup config) — keeps the "." entry
	// byte-stable as the formula entry is added.
	splitting: false,
	// No source maps in the published package (see @openjsxl/core's tsup config).
	sourcemap: false,
	// Keep @openjsxl/core external — the facade re-exports it rather than bundling a second copy.
	// A single copy keeps `instanceof XlsxError` consistent whether a consumer imports from
	// `openjsxl` or `@openjsxl/core`, and aligns the runtime with the type re-export. npm still
	// installs core automatically as a dependency, so `npm i openjsxl` stays one command.
	external: ["@openjsxl/core"],
});
