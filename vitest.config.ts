import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    setupFiles: ['tests/setup-home.ts'],
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: ['src/index.ts'],
      reporter: ['text', 'html'],
    },
  },
});
