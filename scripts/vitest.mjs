// Starts vitest from a working directory with an upper-case drive letter.
// With a lower-case one (d:\..., what Git Bash and VS Code terminals use on Windows) vitest fails with
// "Cannot read properties of undefined (reading 'config')" because it loads itself twice.
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

let root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
if (process.platform === 'win32') root = root.replace(/^[a-z]:/, (d) => d.toUpperCase());
const bin = path.join(root, 'node_modules', 'vitest', 'vitest.mjs');
const res = spawnSync(process.execPath, [bin, ...process.argv.slice(2)], { cwd: root, stdio: 'inherit' });
process.exit(res.status ?? 1);
