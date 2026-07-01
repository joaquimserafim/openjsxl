import { defineConfig } from 'tsup'

export default defineConfig({
	entry: ['src/index.ts'],
	format: ['esm'],
	target: 'node24',
	dts: true,
	clean: true,
	treeshake: true,
	// No source maps in the published package: the output isn't minified (so it's already
	// legible) and the map roughly quadruples install size for little debugging value.
	sourcemap: false,
})
