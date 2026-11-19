// Builds the three bundles (see doc/arch42.md §7):
//   dist/host.js    extension host (node, cjs, own code only)
//   dist/worker.js  Web Worker with the format adapters and third-party libraries
//   dist/webview.js webview UI; embeds worker.js as base64 so no network request or
//                   extra resource is needed to start the worker (connect-src stays 'none')
// plus dist/webview.css (imported from main.ts) and, with --test-build, dist/probe.js.
import * as esbuild from 'esbuild';
import { readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { getBuildInfo } from './scripts/buildinfo.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const buildInfo = getBuildInfo();
const args = new Set(process.argv.slice(2));
const watch = args.has('--watch');
const production = !watch && !args.has('--dev');
const testBuild = args.has('--test-build');

const FORBIDDEN_NODE_MODULES = [
  'http', 'https', 'http2', 'net', 'tls', 'dgram', 'dns', 'child_process',
  'worker_threads', 'cluster', 'fs', 'os', 'vm', 'inspector',
];

/** Fails the build if browser-side code imports a Node module able to do I/O. */
const forbidNodeModules = {
  name: 'forbid-node-modules',
  setup(build) {
    const re = new RegExp(`^(node:)?(${FORBIDDEN_NODE_MODULES.join('|')})(/.*)?$`);
    build.onResolve({ filter: re }, (a) => ({
      errors: [{ text: `Forbidden import of Node module "${a.path}" in browser code (imported from ${a.importer})` }],
    }));
  },
};

/** Provides `virtual:worker-bytes` (dist/worker.js) and `virtual:pdf-worker-bytes` (pdf.js worker) as base64 strings. */
const embedWorker = {
  name: 'embed-worker',
  setup(build) {
    const files = {
      'virtual:worker-bytes': () => path.join(root, 'dist', 'worker.js'),
      'virtual:pdf-worker-bytes': () => path.join(root, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.min.mjs'),
    };
    build.onResolve({ filter: /^virtual:(pdf-)?worker-bytes$/ }, (a) => ({ path: a.path, namespace: 'vsofx-virtual' }));
    build.onLoad({ filter: /.*/, namespace: 'vsofx-virtual' }, async (a) => {
      const file = files[a.path]();
      const b64 = (await readFile(file)).toString('base64');
      return { contents: `export default ${JSON.stringify(b64)};`, loader: 'js', watchFiles: [file] };
    });
  },
};

const common = {
  bundle: true,
  logLevel: 'info',
  sourcemap: production ? false : 'linked',
  minify: production,
  legalComments: 'none',
  metafile: true,
  absWorkingDir: root,
};

const browserCommon = {
  ...common,
  platform: 'browser',
  format: 'iife',
  target: 'es2022',
  plugins: [forbidNodeModules],
  // VS Code's Electron is older than current browsers: pdf.js needs its legacy build (polyfills, e.g. Uint8Array.toHex).
  alias: { 'pdfjs-dist': './node_modules/pdfjs-dist/legacy/build/pdf.mjs' },
  define: { 'process.env.NODE_ENV': production ? '"production"' : '"development"' },
};

const configs = {
  host: {
    ...common,
    entryPoints: ['src/host/extension.ts'],
    outfile: 'dist/host.js',
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['vscode'],
    define: { __BUILD__: JSON.stringify(`${buildInfo.version} (${buildInfo.label})`) },
  },
  worker: {
    ...browserCommon,
    entryPoints: ['src/webview/worker.ts'],
    outfile: 'dist/worker.js',
  },
  webview: {
    ...browserCommon,
    entryPoints: { webview: 'src/webview/main.ts' },
    outdir: 'dist',
    plugins: [forbidNodeModules, embedWorker],
  },
};
if (testBuild) {
  configs.probe = {
    ...browserCommon,
    entryPoints: ['src/webview/probe.ts'],
    outfile: 'dist/probe.js',
  };
  // Integration tests (run inside VS Code by @vscode/test-cli)
  configs.integration = {
    ...common,
    entryPoints: ['test/integration/editor.test.ts'],
    outdir: 'out/test/integration',
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['vscode'],
    sourcemap: 'linked',
    minify: false,
  };
}

async function writeMeta(name, result) {
  if (name === 'integration') return;
  await writeFile(path.join(root, 'dist', `${name}.meta.json`), JSON.stringify(result.metafile));
}

async function run() {
  await rm(path.join(root, 'dist'), { recursive: true, force: true });
  await rm(path.join(root, 'out'), { recursive: true, force: true });
  await mkdir(path.join(root, 'dist'), { recursive: true });
  const order = ['host', 'worker', 'webview', ...(testBuild ? ['probe', 'integration'] : [])];
  if (watch) {
    for (const name of order) {
      const ctx = await esbuild.context({
        ...configs[name],
        plugins: [...(configs[name].plugins ?? []), { name: 'meta', setup: (b) => b.onEnd((r) => r.metafile && writeMeta(name, r)) }],
      });
      await ctx.watch();
    }
    console.log('watching...');
    return;
  }
  for (const name of order) {
    const result = await esbuild.build(configs[name]);
    await writeMeta(name, result);
  }
  const { size } = await import('node:fs').then((fs) => fs.promises.stat(path.join(root, 'dist', 'worker.js')));
  console.log(`worker.js: ${(size / 1024).toFixed(0)} KiB`);
}

run().catch(() => process.exit(1));
