// Entry of dist/worker-sheet.js and dist/worker-doc.js (esbuild.mjs): a classic script that hands the
// source of that family's format worker to main.ts on globalThis. Loaded on demand with the page's nonce.
import source from 'virtual:worker-source';

declare const __FAMILY__: string;

const g = globalThis as { vsofxWorkers?: Record<string, string> };
(g.vsofxWorkers ??= {})[__FAMILY__] = source;
