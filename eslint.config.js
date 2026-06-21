const noUnsanitized = require('eslint-plugin-no-unsanitized');

module.exports = [
  {
    ignores: ['node_modules/**', 'out/**', 'out-publish/**', 'dist/**']
  },
  {
    files: ['**/*.js'],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
      globals: {
        Buffer: 'readonly',
        __dirname: 'readonly',
        clearTimeout: 'readonly',
        console: 'readonly',
        crypto: 'readonly',
        document: 'readonly',
        DOMParser: 'readonly',
        exports: 'readonly',
        module: 'readonly',
        process: 'readonly',
        requestAnimationFrame: 'readonly',
        require: 'readonly',
        Response: 'readonly',
        ResizeObserver: 'readonly',
        setTimeout: 'readonly',
        structuredClone: 'readonly',
        URL: 'readonly',
        URLSearchParams: 'readonly',
        window: 'readonly'
      }
    },
    plugins: {
      'no-unsanitized': noUnsanitized
    },
    rules: {
      'no-constant-condition': ['error', { checkLoops: false }],
      'no-duplicate-imports': 'error',
      'no-undef': 'error',
      'no-unreachable': 'error',
      'no-unused-vars': ['error', { argsIgnorePattern: '^_', caughtErrors: 'none', varsIgnorePattern: '^_' }],
      'no-unsanitized/method': 'error',
      'no-unsanitized/property': 'error'
    }
  },
  {
    files: ['main.js', 'preload.js', 'forge.config.js', 'eslint.config.js', 'test/**/*.js'],
    languageOptions: {
      sourceType: 'commonjs'
    },
    rules: {
      'no-unsanitized/method': 'off'
    }
  }
];
