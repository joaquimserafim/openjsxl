import { defineConfig } from "tsup";

export default defineConfig({
	// Two entries: the always-loaded core (".") and the opt-in formula parser ("./formula", M8).
	entry: { index: "src/index.ts", formula: "src/formula/index.ts" },
	format: ["esm"],
	target: "node24",
	dts: true,
	clean: true,
	treeshake: true,
	// Emit each entry as one self-contained file. A second ESM entry would otherwise let esbuild
	// hoist shared helpers into a chunk and churn dist/index.js — breaking the "." entry's
	// byte-identity. With splitting off the formula bundle duplicates the few helpers it shares (a
	// few KB, accepted) and dist/index.js stays exactly what it was before this entry existed.
	splitting: false,
	// No source maps in the published package: the output isn't minified (so it's already
	// legible) and the map roughly quadruples install size for little debugging value.
	sourcemap: false,
});
