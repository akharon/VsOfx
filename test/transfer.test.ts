// P4 (bytes as typed arrays, transferred) and U1 (loading skeleton in the page HTML).
import { afterEach, describe, expect, it } from 'vitest';
import { asBytes, ownedBytes } from '../src/shared/bytes.ts';
import { fromByteArray } from '../src/shared/base64.ts';
import { skeletonHtml, skeletonKind } from '../src/shared/skeleton.ts';
import { buildCsp, buildWebviewHtml } from '../src/host/webviewHtml.ts';
import { workerFamily } from '../src/webview/family.ts';
import { runWorker } from '../src/webview/workers/core.ts';
import { CsvAdapter } from '../src/webview/adapters/csvAdapter.ts';
import type { WorkerRequest, WorkerResponse } from '../src/webview/protocol.ts';
import { readFileSync } from 'node:fs';

describe('asBytes / ownedBytes', () => {
  const sample = Uint8Array.from([1, 2, 3, 250]);

  it('accepts typed arrays, ArrayBuffers and views', () => {
    expect(asBytes(sample)).toBe(sample);
    expect(Array.from(asBytes(sample.buffer))).toEqual([1, 2, 3, 250]);
    expect(Array.from(asBytes(new DataView(sample.buffer, 1, 2)))).toEqual([2, 3]);
    expect(Array.from(asBytes(Buffer.from([7, 8])))).toEqual([7, 8]);
  });

  it('still accepts a serialised Buffer, a number array and the old base64 string', () => {
    expect(Array.from(asBytes({ type: 'Buffer', data: [4, 5] }))).toEqual([4, 5]);
    expect(Array.from(asBytes([9, 10]))).toEqual([9, 10]);
    expect(Array.from(asBytes(fromByteArray(sample)))).toEqual([1, 2, 3, 250]);
  });

  it('rejects anything else', () => {
    expect(() => asBytes(undefined)).toThrow(/file bytes/);
    expect(() => asBytes({ nope: 1 })).toThrow(/file bytes/);
    expect(() => asBytes(5)).toThrow(/file bytes/);
  });

  it('ownedBytes gives a buffer that can be transferred whole', () => {
    const big = new Uint8Array([0, 1, 2, 3, 4, 5]);
    const view = big.subarray(2, 4);
    const own = ownedBytes(view);
    expect(Array.from(own)).toEqual([2, 3]);
    expect(own.buffer.byteLength).toBe(2);
    expect(ownedBytes(big)).toBe(big);
  });
});

describe('worker byte transfer', () => {
  const g = globalThis as { self?: unknown };
  const saved = g.self;
  afterEach(() => {
    g.self = saved;
  });

  it('hands the saved bytes over as a transferable and stays usable for the next save', async () => {
    const listeners: ((ev: { data: WorkerRequest }) => void)[] = [];
    const replies: { msg: WorkerResponse; transfer?: Transferable[] }[] = [];
    g.self = {
      addEventListener: (_: string, fn: (ev: { data: WorkerRequest }) => void) => listeners.push(fn),
      postMessage: (msg: WorkerResponse, transfer?: Transferable[]) => {
        // really detach what is transferred, like a browser does
        const copy = structuredClone(msg, { transfer });
        replies.push({ msg: copy, transfer });
      },
    };
    runWorker(() => new CsvAdapter());
    const send = async (req: WorkerRequest) => {
      const before = replies.length;
      for (const l of listeners) l({ data: req });
      for (let i = 0; i < 50 && replies.length === before; i++) await new Promise((r) => setTimeout(r, 5));
      return replies[replies.length - 1]!;
    };
    const text = new TextEncoder().encode('a,b\n1,2\n');
    const loaded = await send({ id: 1, op: 'load', format: 'csv', bytes: text, edits: [] });
    expect(loaded.msg.ok).toBe(true);
    for (const id of [2, 3]) {
      const r = await send({ id, op: 'save' });
      expect(r.msg.ok).toBe(true);
      const out = (r.msg as { result: Uint8Array }).result;
      expect(out).toBeInstanceOf(Uint8Array);
      expect(new TextDecoder().decode(out)).toBe('a,b\n1,2\n');
      expect(r.transfer).toHaveLength(1);
    }
  });
});

describe('format workers per family', () => {
  it('maps every format to a family', () => {
    expect(['xlsx', 'ods', 'csv'].map((f) => workerFamily(f as never))).toEqual(['sheet', 'sheet', 'sheet']);
    expect(['docx', 'odt', 'pptx', 'odp', 'odg'].map((f) => workerFamily(f as never))).toEqual(['doc', 'doc', 'doc', 'doc', 'doc']);
    expect(workerFamily('pdf')).toBe('pdf');
  });
});

describe('loading skeleton', () => {
  const base = { cspSource: 'vscode-resource:', scriptUri: 'vscode-resource:/webview.js', styleUri: 'vscode-resource:/webview.css', nonce: 'abc' };

  it('is in the page HTML before any script runs, for the format at hand', () => {
    expect(buildWebviewHtml({ ...base, format: 'xlsx' })).toContain('class="skeleton sk-grid"');
    expect(buildWebviewHtml({ ...base, format: 'csv' })).toContain('data-skeleton="grid"');
    expect(buildWebviewHtml({ ...base, format: 'docx' })).toContain('class="skeleton sk-page"');
    expect(buildWebviewHtml(base)).toContain('class="skeleton sk-page"');
    expect(skeletonKind('pdf')).toBe('page');
  });

  it('is static markup: no scripts, handlers or URLs, and the policy does not change', () => {
    for (const k of ['grid', 'page'] as const) {
      const html = skeletonHtml(k);
      expect(html).not.toMatch(/<script|\son\w+=|https?:|src=|href=/i);
      expect(html).toContain('role="status"');
      expect(html).toContain('Loading…');
    }
    const html = buildWebviewHtml({ ...base, format: 'ods' });
    expect(html).toContain(`content="${buildCsp('vscode-resource:', 'abc')}"`);
    expect(html.match(/<script[^>]*>/g)).toHaveLength(1);
  });

  it('is styled with theme variables and stands still for reduced motion', () => {
    const css = readFileSync(new URL('../src/webview/style.css', import.meta.url), 'utf8');
    expect(css).toContain('.sk-toolbar');
    expect(css).toMatch(/@media \(prefers-reduced-motion: reduce\)\s*\{[^}]*\.sk-/);
    expect(css).toMatch(/@keyframes sk-shimmer/);
    const block = css.slice(css.indexOf('loading skeleton'));
    expect(block).toContain('var(--vscode-');
    expect(block).not.toMatch(/#[0-9a-f]{3,6}\b/i);
  });
});
