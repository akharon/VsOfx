// Smoke test of the real dist/worker-<family>.raw.js (the browser bundles that ship, embedded in dist/worker-*.js and dist/pdf.js), run in a bare `vm`
// context that has browser-like globals only: no process, require, Buffer. Network entry points are
// replaced by throwing functions and must never be called. Needs `npm run build` first.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { describe, expect, it } from 'vitest';
import type { DocModel, Edit, Format } from '../src/shared/model.ts';
import { XlsxAdapter } from '../src/webview/adapters/xlsxAdapter.ts';
import { DocxAdapter } from '../src/webview/adapters/docxAdapter.ts';
import { isZip } from '../src/webview/adapters/zip.ts';
import { makeDocx, makeOds, makeOdt, makePptx, makeXlsx } from './helpers/fixtures.ts';

import { workerFamily, type WorkerFamily } from '../src/webview/family.ts';

const dist = path.join(__dirname, '..', 'dist');
const bundleOf = (family: WorkerFamily) => path.join(dist, `worker-${family}.raw.js`);
const built = (['sheet', 'doc', 'pdf'] as const).every((f) => fs.existsSync(bundleOf(f)));

interface Reply {
  id: number;
  ok: boolean;
  message?: string;
  result?: unknown;
}

function startWorker(family: WorkerFamily) {
  const listeners: ((ev: { data: unknown }) => void)[] = [];
  const replies = new Map<number, (r: Reply) => void>();
  const networkAttempts: string[] = [];
  const deny = (name: string) =>
    function () {
      networkAttempts.push(name);
      throw new Error(`network attempt: ${name}`);
    };
  const sandbox: Record<string, unknown> = {
    console,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    queueMicrotask,
    structuredClone,
    MessageChannel,
    TextEncoder,
    TextDecoder,
    URL,
    atob,
    btoa,
    performance,
    crypto: globalThis.crypto,
    navigator: { userAgent: 'vm-worker' },
    fetch: deny('fetch'),
    XMLHttpRequest: deny('XMLHttpRequest'),
    WebSocket: deny('WebSocket'),
    EventSource: deny('EventSource'),
    importScripts: deny('importScripts'),
    addEventListener: (_: string, fn: (ev: { data: unknown }) => void) => listeners.push(fn),
    postMessage: (msg: Reply) => replies.get(msg.id)?.(msg),
  };
  sandbox['self'] = sandbox;
  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(bundleOf(family), 'utf8'), context, { filename: `worker-${family}.js` });
  let nextId = 1;

  function request(req: Record<string, unknown>, bytes?: Uint8Array): Promise<Reply> {
    const id = nextId++;
    sandbox['__json'] = JSON.stringify({ ...req, id });
    sandbox['__b64'] = bytes ? Buffer.from(bytes).toString('base64') : '';
    const msg = vm.runInContext(
      `(() => { const m = JSON.parse(__json); if (__b64) m.bytes = Uint8Array.from(atob(__b64), c => c.charCodeAt(0)); return m; })()`,
      context,
    );
    return new Promise((resolve) => {
      replies.set(id, resolve);
      for (const l of listeners) l({ data: msg });
    });
  }
  return {
    request,
    networkAttempts,
    globals: () => vm.runInContext('typeof process + "," + typeof require + "," + typeof Buffer + "," + typeof window + "," + typeof document', context) as string,
    load: (format: Format, bytes: Uint8Array, edits: Edit[] = []) => request({ op: 'load', format, edits }, bytes),
    apply: (edit: Edit) => request({ op: 'apply', edit }),
    save: async () => {
      const r = await request({ op: 'save' });
      expect(r.ok, r.message).toBe(true);
      return Uint8Array.from(r.result as ArrayLike<number>);
    },
  };
}

describe.skipIf(!built)('format workers in a browser-like sandbox', () => {
  it('run without Node globals', () => {
    for (const f of ['sheet', 'doc', 'pdf'] as const) {
      expect(startWorker(f).globals()).toBe('undefined,undefined,undefined,undefined,undefined');
    }
  });

  it('each family bundle holds only its own libraries', () => {
    const libs = (f: WorkerFamily) => {
      const meta = JSON.parse(fs.readFileSync(path.join(dist, `worker-${f}.meta.json`), 'utf8')) as { inputs: Record<string, unknown> };
      const has = (lib: string) => Object.keys(meta.inputs).some((p) => p.includes(`node_modules/${lib}/`));
      return { exceljs: has('exceljs'), mammoth: has('mammoth'), pdfLib: has('@cantoo/pdf-lib') };
    };
    expect(libs('sheet')).toEqual({ exceljs: true, mammoth: false, pdfLib: false });
    expect(libs('doc')).toEqual({ exceljs: false, mammoth: true, pdfLib: false });
    expect(libs('pdf')).toEqual({ exceljs: false, mammoth: false, pdfLib: true });
    expect(fs.statSync(bundleOf('doc')).size).toBeLessThan(fs.statSync(bundleOf('sheet')).size / 2);
  });

  it('rejects a format of another family', async () => {
    const r = await startWorker('doc').load('xlsx', await makeXlsx('styles'));
    expect(r.ok).toBe(false);
    expect(r.message).toMatch(/cannot open xlsx/);
  });

  it('xlsx: load, edit, save with exceljs', async () => {
    const w = startWorker('sheet');
    const loaded = await w.load('xlsx', await makeXlsx('styles'));
    expect(loaded.ok, loaded.message).toBe(true);
    const model = (loaded.result as { model: DocModel }).model;
    expect(model.type).toBe('sheet');
    expect(JSON.stringify(model)).toContain('1,234.50');
    const applied = await w.apply({ kind: 'cell', sheet: 0, row: 1, col: 1, value: '77', old: '1234.5' });
    expect(applied.ok, applied.message).toBe(true);
    const saved = await w.save();
    expect(isZip(saved)).toBe(true);
    const back = (await new XlsxAdapter().load(saved, [])) as Extract<DocModel, { type: 'sheet' }>;
    expect(back.sheets[0]!.rows[1]![1]!.text).toBe('77.00');
    expect(w.networkAttempts).toEqual([]);
  });

  it('docx: load with mammoth, edit, save', async () => {
    const w = startWorker('doc');
    const loaded = await w.load('docx', makeDocx('styles'));
    expect(loaded.ok, loaded.message).toBe(true);
    const html = (loaded.result as { model: { html: string } }).model.html;
    expect(html).toContain('<h1 data-p="0">Main Title</h1>');
    await w.apply({ kind: 'paragraph', index: 1, text: 'Edited in worker', old: 'Intro text.' });
    const back = (await new DocxAdapter().load(await w.save(), [])) as Extract<DocModel, { type: 'text' }>;
    expect(back.html).toContain('Edited in worker');
    expect(w.networkAttempts).toEqual([]);
  });

  it('pptx, odt and ods load', async () => {
    expect(workerFamily('ods')).toBe('sheet');
    const w = startWorker('doc');
    const p = await w.load('pptx', makePptx());
    expect((p.result as { model: { slides: unknown[] }; readonly: boolean }).model.slides).toHaveLength(2);
    expect((p.result as { readonly: boolean }).readonly).toBe(true);
    const t = await w.load('odt', makeOdt());
    expect((t.result as { model: { html: string } }).model.html).toContain('Report title');
    const s = await startWorker('sheet').load('ods', makeOds());
    expect((s.result as { model: { sheets: unknown[] } }).model.sheets).toHaveLength(2);
    expect(w.networkAttempts).toEqual([]);
  });

  it('reports a clear error for files that are not Office packages', async () => {
    const w = startWorker('doc');
    const r = await w.load('docx', new TextEncoder().encode('plain text, not a zip'));
    expect(r.ok).toBe(false);
    expect(r.message).toMatch(/not a valid Office document/);
    const r2 = await w.apply({ kind: 'paragraph', index: 0, text: 'x', old: '' });
    expect(r2.ok).toBe(false);
  });
});
