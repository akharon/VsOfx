// Build checks for the "no network" guarantee (doc/arch42.md §8.1 item 6).
//  1. The host bundle contains no third-party code (no node_modules inputs).
//  2. dist/*.js is scanned for network / dynamic-code primitives. Every hit must be
//     listed (file, pattern, count, reason) in scripts/allowed-hits.json; new hits fail.
//     The allowlist is reviewed by hand: the Content-Security-Policy is the enforcement,
//     this check makes any change in third-party behaviour visible.
//  Use --update to print a fresh allowlist skeleton.
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const dist = path.join(root, 'dist');
const PATTERNS = [
  ['fetch(', /\bfetch\s*\(/g],
  ['XMLHttpRequest', /XMLHttpRequest/g],
  ['WebSocket', /WebSocket/g],
  ['EventSource', /EventSource/g],
  ['sendBeacon', /sendBeacon/g],
  ['importScripts', /importScripts/g],
  ['eval(', /\beval\s*\(/g],
  ['new Function', /new\s+Function\b|\bFunction\s*\(\s*["'`]/g],
];
const SCANNED = ['host.js', 'worker-sheet.js', 'worker-doc.js', 'pdf.js', 'webview.js']; // probe.js is test-only and never shipped
const update = process.argv.includes('--update');

let failed = false;
const fail = (msg) => {
  failed = true;
  console.error(`check-bundles: ${msg}`);
};

// 1. host bundle must not include third-party code
try {
  const meta = JSON.parse(await readFile(path.join(dist, 'host.meta.json'), 'utf8'));
  const bad = Object.keys(meta.inputs).filter((p) => p.includes('node_modules'));
  if (bad.length) fail(`host bundle contains third-party code: ${bad.slice(0, 5).join(', ')}`);
} catch (e) {
  fail(`cannot read dist/host.meta.json (run npm run build first): ${e.message}`);
}

// 2. pattern scan
const allowed = JSON.parse(await readFile(path.join(root, 'scripts', 'allowed-hits.json'), 'utf8')).hits;
const found = [];
for (const file of SCANNED) {
  let code;
  try {
    code = await readFile(path.join(dist, file), 'utf8');
  } catch {
    fail(`dist/${file} is missing (run npm run build first)`);
    continue;
  }
  for (const [name, re] of PATTERNS) {
    const count = (code.match(re) ?? []).length;
    if (count > 0) found.push({ file, pattern: name, count });
  }
}

if (update) {
  console.log(JSON.stringify({ hits: found.map((h) => ({ ...h, reason: 'TODO' })) }, null, 2));
  process.exit(0);
}

for (const h of found) {
  const a = allowed.find((x) => x.file === h.file && x.pattern === h.pattern);
  if (!a) fail(`${h.file}: ${h.count} new hit(s) of "${h.pattern}" – review and add to scripts/allowed-hits.json with a reason`);
  else if (h.count > a.count) fail(`${h.file}: "${h.pattern}" appears ${h.count} times, allowlist says ${a.count} – review the new occurrences`);
}
for (const a of allowed) {
  if (!found.some((h) => h.file === a.file && h.pattern === a.pattern)) {
    console.warn(`check-bundles: allowlist entry no longer needed: ${a.file} "${a.pattern}"`);
  }
}

if (failed) process.exit(1);
console.log(`check-bundles: OK (${found.length} reviewed hit(s) in ${SCANNED.length} bundles)`);
