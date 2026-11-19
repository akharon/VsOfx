// Generates THIRD_PARTY_NOTICES.md for the production dependency tree and fails on licenses
// outside the allowlist (doc/plan.md 0.6). `--check` only verifies, `npm run notices` also writes the file.
import { writeFile, readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const checker = require('license-checker-rseidelsohn');
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const checkOnly = process.argv.includes('--check');

const ALLOWED = new Set(['MIT', 'ISC', 'BSD-2-Clause', 'BSD-3-Clause', 'Apache-2.0', '0BSD', 'Unlicense', 'CC0-1.0', 'BlueOak-1.0.0', 'Python-2.0']);
// Packages whose license metadata is unusual but was reviewed by hand.
const REVIEWED = {
  'buffers@0.1.1': 'no license metadata; upstream (substack/node-buffers) is MIT/X11; Node-only transitive dependency of exceljs streaming (unzipper), not in any shipped bundle',
  'duck@0.1.12': 'BSD-style license file present; dependency of mammoth (lop)',
  'pako@1.0.11': '(MIT AND Zlib): both permissive',
};

const packages = await new Promise((resolve, reject) =>
  checker.init({ start: root, production: true, excludePrivatePackages: true }, (err, json) => (err ? reject(err) : resolve(json))),
);

/** "(MIT OR Apache-2.0)" is fine if at least one alternative is allowed; "AND" needs all. */
function acceptable(expr) {
  const clean = expr.replace(/[()*]/g, ' ').trim();
  if (/\bOR\b/.test(clean)) return clean.split(/\bOR\b/).some((p) => acceptable(p));
  return clean.split(/\bAND\b/).every((p) => ALLOWED.has(p.trim()));
}

const bad = [];
let out = '# Third-party notices\n\nVsOfx bundles the following third-party software (production dependency tree).\n';
for (const [id, info] of Object.entries(packages).sort(([a], [b]) => a.localeCompare(b))) {
  const license = String(info.licenses);
  if (!REVIEWED[id] && !acceptable(license)) bad.push(`${id}: ${license}`);
  let text = '';
  if (info.licenseFile) {
    try {
      text = (await readFile(info.licenseFile, 'utf8')).trim();
    } catch {
      /* no file */
    }
  }
  out += `\n---\n\n## ${id}\n\nLicense: ${license}\n${info.repository ? `Source: ${info.repository}\n` : ''}\n`;
  if (text && !/readme/i.test(path.basename(info.licenseFile))) out += '```\n' + text + '\n```\n';
}
out += '\n---\n\n## Microsoft vscode-extension-samples (custom-editor-sample)\n\nLicense: MIT. The structure of the custom editor provider follows this sample. Copyright (c) Microsoft Corporation.\n';

if (bad.length) {
  console.error('notices: licenses outside the allowlist:\n  ' + bad.join('\n  '));
  process.exit(1);
}
if (checkOnly) {
  console.log(`notices: OK (${Object.keys(packages).length} packages, all licenses allowed)`);
} else {
  await writeFile(path.join(root, 'THIRD_PARTY_NOTICES.md'), out);
  console.log(`notices: wrote THIRD_PARTY_NOTICES.md (${Object.keys(packages).length} packages)`);
}
