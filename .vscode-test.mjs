import { defineConfig } from '@vscode/test-cli';

// Integration tests run inside a real VS Code. Set VSCODE_EXECUTABLE_PATH to use an installed
// VS Code instead of downloading one. Build first: npm run test:integration does it.
export default defineConfig({
  files: 'out/test/integration/**/*.test.js',
  ...(process.env.VSCODE_EXECUTABLE_PATH ? { useInstallation: { fromPath: process.env.VSCODE_EXECUTABLE_PATH } } : { version: 'stable' }),
  launchArgs: ['--disable-extensions', '--disable-workspace-trust'],
  mocha: { ui: 'bdd', timeout: 90_000 },
});
