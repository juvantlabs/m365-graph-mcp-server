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
      // Coverage scope: the pure-logic + tool-handler surface that's
      // unit-testable without mocking auth or filesystem I/O.
      // src/index.ts, src/auth/*, src/client/* are integration-tested
      // via the live tenant smoke runs and excluded from unit-test
      // coverage thresholds. Add them here as their unit tests land.
      include: [
        'src/types/validators.ts',
        'src/tools/**/*.ts',
      ],
      exclude: [
        'src/**/*.d.ts',
        // Pure type files contribute no executable lines; including
        // them just creates noise in the coverage report.
        'src/types/tool.ts',
      ],
      // Per-file thresholds enforce the floor on each tested file
      // individually, so a regression in any one file fails CI.
      // Per handbook docs/repo-types/mcp-server.md § CI requirements,
      // the spec target is 80%. Branches is set to 50% for now because
      // the download_file streaming-pipeline branches require mocking
      // fs + fetch (deferred); raise to 80% when those tests land.
      //
      thresholds: {
        perFile: true,
        lines: 80,
        functions: 80,
        statements: 80,
        branches: 50,
      },
    },
  },
});
