// Builds the bundles (see doc/arch42.md §7):
//   dist/host.js    extension host (node, cjs, own code only)
//   dist/worker-<family>.raw.js  Web Worker with the adapters and libraries of one family (sheet: xlsx, ods,
//                   csv; doc: docx, odt, pptx, odp; pdf: pdf-lib). Intermediate, not shipped.
//   dist/worker-sheet.js, worker-doc.js  the worker source as a JS string literal in a classic script that
//                   sets globalThis.vsofxWorkers[family]; loaded by webview.js when such a file is opened
//   dist/webview.js webview UI (small: no libraries)
//   dist/pdf.js     pdf.js + PDF view + the pdf.js worker source + the pdf format worker source (string
//                   literals); a classic script that sets globalThis.vsofxPdf, loaded when a PDF is opened
// plus dist/webview.css (imported from main.ts) and, with --test-build, dist/probe.js.
import * as esbuild from 'esbuild';
import { readFile, writeFile, mkdir, rm, stat } from 'node:fs/promises';
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

/**
 * Provides `virtual:worker-source` (dist/worker-<family>.raw.js, the format worker of one family) and
 * `virtual:pdf-worker-source` (pdf.js worker) as JS string literals.
 */
const embedWorker = (family) => ({
  name: 'embed-worker',
  setup(build) {
    const files = {
      'virtual:worker-source': () => path.join(root, 'dist', `worker-${family}.raw.js`),
      'virtual:pdf-worker-source': () => path.join(root, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.min.mjs'),
    };
    build.onResolve({ filter: /^virtual:(pdf-)?worker-source$/ }, (a) => ({ path: a.path, namespace: 'vsofx-virtual' }));
    build.onLoad({ filter: /.*/, namespace: 'vsofx-virtual' }, async (a) => {
      const file = files[a.path]();
      const text = await readFile(file, 'utf8');
      return { contents: `export default ${JSON.stringify(text)};`, loader: 'js', watchFiles: [file] };
    });
  },
});

/** Format workers, one per family, so a document does not parse the libraries of the other families. */
const FAMILIES = ['sheet', 'doc', 'pdf'];

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
  webview: {
    ...browserCommon,
    entryPoints: { webview: 'src/webview/main.ts' },
    outdir: 'dist',
  },
  pdf: {
    ...browserCommon,
    entryPoints: ['src/webview/pdfChunk.ts'],
    outfile: 'dist/pdf.js',
    plugins: [forbidNodeModules, embedWorker('pdf')],
  },
};
for (const f of FAMILIES) {
  // the worker itself (not shipped) ...
  configs[`worker-${f}`] = { ...browserCommon, entryPoints: [`src/webview/workers/${f}.ts`], outfile: `dist/worker-${f}.raw.js` };
  // ... and the script that hands its source over on globalThis (shipped, loaded on demand). The PDF
  // worker source travels inside dist/pdf.js instead.
  if (f !== 'pdf') {
    configs[`chunk-${f}`] = {
      ...browserCommon,
      entryPoints: ['src/webview/workerChunk.ts'],
      outfile: `dist/worker-${f}.js`,
      plugins: [forbidNodeModules, embedWorker(f)],
      define: { ...browserCommon.define, __FAMILY__: JSON.stringify(f) },
    };
  }
}
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
  const order = [
    'host',
    ...FAMILIES.map((f) => `worker-${f}`),
    ...FAMILIES.filter((f) => f !== 'pdf').map((f) => `chunk-${f}`),
    'pdf',
    'webview',
    ...(testBuild ? ['probe', 'integration'] : []),
  ];
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
  for (const f of ['host.js', 'worker-sheet.js', 'worker-doc.js', 'pdf.js', 'webview.js']) {
    const { size } = await stat(path.join(root, 'dist', f));
    console.log(`${f}: ${(size / 1024).toFixed(0)} KiB`);
  }
}

run().catch(() => process.exit(1));
