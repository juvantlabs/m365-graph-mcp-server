// ESLint flat config (v9+) for the MCP server.
//
// Stdout discipline: console.log is forbidden anywhere in src/ — it
// corrupts the MCP stdio JSON-RPC framing. Use console.error for
// diagnostics. This rule is also enforced by a CI grep step in
// .github/workflows/ci.yml as defense-in-depth.

import tsParser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';

const sharedRules = {
  // Stdout discipline — non-negotiable for MCP stdio servers.
  'no-console': ['error', { allow: ['error', 'warn'] }],

  // Strict TypeScript hygiene.
  '@typescript-eslint/no-explicit-any': 'error',
  '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
  '@typescript-eslint/explicit-function-return-type': ['warn', {
    allowExpressions: true,
    allowTypedFunctionExpressions: true,
  }],

  // General hygiene.
  'eqeqeq': ['error', 'always'],
  'no-throw-literal': 'error',
  'prefer-const': 'error',
};

export default [
  {
    ignores: ['dist/**', 'node_modules/**', 'coverage/**'],
  },
  {
    // Production source: full type-aware linting via tsconfig.
    files: ['src/**/*.ts'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
        project: './tsconfig.json',
      },
    },
    plugins: { '@typescript-eslint': tsPlugin },
    rules: {
      ...sharedRules,
      // Type-aware rules — only valid when parserOptions.project is set.
      '@typescript-eslint/no-floating-promises': 'error',
    },
  },
  {
    // Test files: same hygiene rules but no type-aware project linkage,
    // so tests aren't required to live inside the production tsconfig
    // include. tsc still typechecks them via npm run typecheck if they
    // get added to a tsconfig include later.
    files: ['tests/**/*.ts'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
      },
    },
    plugins: { '@typescript-eslint': tsPlugin },
    rules: sharedRules,
  },
];
