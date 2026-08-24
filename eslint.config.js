import js from '@eslint/js'
import tseslint from 'typescript-eslint'
import react from 'eslint-plugin-react'
import reactHooks from 'eslint-plugin-react-hooks'
import globals from 'globals'

export default tseslint.config(
  {
    ignores: [
      'node_modules/**',
      'dist/**',
      'build/**',
      'packages/*/dist/**',
      'packages/*/node_modules/**',
      'coverage/**',
      'spike/**',
      'installer/staging/**',
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      // `const { omitMe, ...rest } = obj` is the idiomatic way to drop a key,
      // and a leading underscore is the conventional marker for a binding that
      // exists only to be discarded. Neither is a real unused variable.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          ignoreRestSiblings: true,
        },
      ],
    },
  },
  {
    // Every package that renders React: the shared UI and both shells.
    files: ['packages/{ui,addin,extension}/**/*.{ts,tsx}'],
    languageOptions: {
      parserOptions: {
        ecmaFeatures: { jsx: true },
      },
      globals: { ...globals.browser },
    },
    plugins: { react, 'react-hooks': reactHooks },
    settings: { react: { version: '18' } },
    rules: {
      ...react.configs.recommended.rules,
      ...reactHooks.configs.recommended.rules,
      'react/react-in-jsx-scope': 'off',
      'react/prop-types': 'off',
    },
  },
  {
    files: ['packages/host/**/*.ts'],
    languageOptions: {
      globals: { ...globals.node },
    },
  },
  {
    // Build-time scripts. They run under Node, not in a pane or a panel.
    files: ['scripts/**/*.{js,mjs,cjs}', 'packages/*/scripts/**/*.{js,mjs,cjs}'],
    languageOptions: {
      globals: { ...globals.node },
    },
  },
  {
    // The extension runs against the Chrome APIs, which are ambient globals.
    files: ['packages/extension/**/*.{ts,tsx}'],
    languageOptions: {
      globals: { ...globals.browser, chrome: 'readonly' },
    },
  },
  {
    files: ['packages/shared/**/*.ts'],
    languageOptions: {
      globals: { ...globals.node, ...globals.browser },
    },
  },
)
