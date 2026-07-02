import { defineConfig } from "tsup"

export default defineConfig({
	entry: ["src/index.ts"],
	format: ["esm"],
	target: "node24",
	dts: true,
	clean: true,
	treeshake: true,
	// No source maps in the published package (see @openjsxl/core's tsup config).
	sourcemap: false,
	// Keep @openjsxl/core external — the facade re-exports it rather than bundling a second copy.
	// A single copy keeps `instanceof XlsxError` consistent whether a consumer imports from
	// `openjsxl` or `@openjsxl/core`, and aligns the runtime with the type re-export. npm still
	// installs core automatically as a dependency, so `npm i openjsxl` stays one command.
	external: ["@openjsxl/core"],
})
