import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'node',
    include: ['test/**/*.test.ts'],
    exclude: ['test/integration/**'],
    // Third-party code must never reach the network: every test runs with the
    // network globals replaced by functions that throw (see test/setup.ts).
    setupFiles: ['test/setup.ts'],
    testTimeout: 30000,
  },
});
