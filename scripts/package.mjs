// npm run package: build, notices, then a VSIX whose version and name come from git (scripts/buildinfo.mjs).
// package.json is not modified; vsce writes the version only into the packaged manifest.
import { spawnSync } from 'node:child_process';
import { statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { getBuildInfo } from './buildinfo.mjs';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const info = getBuildInfo();
const out = `vsofx-${info.version}-${info.label}.vsix`;
const run = (cmd, args) => {
  const r = spawnSync(cmd, args, { cwd: root, stdio: 'inherit', shell: true });
  if (r.status !== 0) process.exit(r.status ?? 1);
};
console.log(`Packaging ${out} (version ${info.version}, commit ${info.label})`);
run('npm', ['run', 'build']);
run('npm', ['run', 'notices']);
run('npx', ['vsce', 'package', info.version, '--no-update-package-json', '--no-git-tag-version', '--no-dependencies', '--out', out]);

// Size budget for the .vsix.
const MAX_VSIX_BYTES = 1_992_294; // 1.9 MiB
const { size } = statSync(path.join(root, out));
console.log(`${out}: ${(size / 1024).toFixed(0)} KB (limit ${(MAX_VSIX_BYTES / 1024).toFixed(0)} KB)`);
if (size > MAX_VSIX_BYTES) {
  console.error(`package: ${out} is ${size} bytes, larger than the budget of ${MAX_VSIX_BYTES} bytes`);
  process.exit(1);
}
