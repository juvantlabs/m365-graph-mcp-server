import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Default: unit tests only. Override with `vitest run tests/integration`
    // or via the test:integration npm script.
    include: ['tests/unit/**/*.test.ts'],
    environment: 'node',
    globals: false,
    testTimeout: 10_000,
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.d.ts', 'src/types/**'],
      thresholds: {
        // Per handbook docs/repo-types/mcp-server.md § CI requirements: ≥ 80%.
        lines: 80,
        functions: 80,
        branches: 80,
        statements: 80,
      },
    },
  },
});
