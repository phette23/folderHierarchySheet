import js from '@eslint/js'
import globals from 'globals'

export default [
  { ignores: ['node_modules/**'] },
  js.configs.recommended,
  {
    files: ['**/*.js', '**/*.mjs'],
    languageOptions: { ecmaVersion: 2022, globals: globals.node },
    rules: { semi: ['error', 'never'] },
  },
  {
    files: ['src/*.js'],
    languageOptions: {
      sourceType: 'script',
      globals: { DriveApp: 'readonly', SpreadsheetApp: 'readonly', LockService: 'readonly' },
    },
    rules: { 'no-unused-vars': ['error', { varsIgnorePattern: '^(onOpen|promptFolderHierarchy|createFolderHierarchy)$' }] },
  },
]
