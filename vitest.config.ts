import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

const resolvePath = (p: string) => fileURLToPath(new URL(p, import.meta.url))

export default defineConfig({
  resolve: {
    // Run tests against TypeScript source — no build step required in dev.
    alias: {
      '@openjsxl/core': resolvePath('./packages/core/src/index.ts'),
      openjsxl: resolvePath('./packages/openjsxl/src/index.ts'),
    },
  },
  test: {
    include: ['packages/*/src/**/*.test.ts', 'packages/*/test/**/*.test.ts'],
  },
})
