// ESLint flat config (v9+) for the MCP server.
//
// Stdout discipline: console.log is forbidden anywhere in src/ — it
// corrupts the MCP stdio JSON-RPC framing. Use console.error for
// diagnostics. This rule is also enforced by a CI grep step in
// .github/workflows/ci.yml as defense-in-depth.

import tsParser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';

export default [
  {
    ignores: ['dist/**', 'node_modules/**', 'coverage/**'],
  },
  {
    files: ['src/**/*.ts', 'tests/**/*.ts'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
        project: './tsconfig.json',
      },
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
    },
    rules: {
      // Stdout discipline — non-negotiable for MCP stdio servers.
      'no-console': ['error', { allow: ['error', 'warn'] }],

      // Strict TypeScript hygiene.
      '@typescript-eslint/no-explicit-any': 'error',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/explicit-function-return-type': ['warn', {
        allowExpressions: true,
        allowTypedFunctionExpressions: true,
      }],
      '@typescript-eslint/no-floating-promises': 'error',

      // General hygiene.
      'eqeqeq': ['error', 'always'],
      'no-throw-literal': 'error',
      'prefer-const': 'error',
    },
  },
];
