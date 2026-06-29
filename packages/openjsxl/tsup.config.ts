import { defineConfig } from 'tsup'

export default defineConfig({
	entry: ['src/index.ts'],
	format: ['esm'],
	target: 'node24',
	dts: true,
	clean: true,
	treeshake: true,
	sourcemap: true,
	// Bundle the core engine so consumers install a single package.
	noExternal: ['@openjsxl/core'],
})
