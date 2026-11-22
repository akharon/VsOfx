// Web Worker body: all third-party parsing happens here (doc/arch42.md AD-7). Each family entry
// (sheet.ts, doc.ts, pdf.ts) passes the factory for its own formats, so a worker only contains the
// libraries it needs.
import type { FormatAdapter } from '../adapters/Adapter';
import { assertOfficePackage, assertPdf, withBatches } from '../adapters/common';
import { isTextFormat, type Format } from '../../shared/model';
import type { LoadResult, WorkerRequest, WorkerResponse } from '../protocol';

export function runWorker(factory: (format: Format) => FormatAdapter): void {
  let adapter: FormatAdapter | undefined;

  const scope = self as unknown as {
    addEventListener(type: 'message', fn: (ev: { data: WorkerRequest }) => void): void;
    postMessage(msg: WorkerResponse, transfer?: Transferable[]): void;
  };

  scope.addEventListener('message', (ev) => {
    void handle(ev.data);
  });

  async function handle(req: WorkerRequest): Promise<void> {
    try {
      switch (req.op) {
        case 'load': {
          if (req.format === 'pdf') assertPdf(req.bytes);
          else if (!isTextFormat(req.format)) assertOfficePackage(req.bytes);
          adapter = withBatches(factory(req.format));
          const model = await adapter.load(req.bytes, req.edits);
          const result: LoadResult = { model, readonly: adapter.readonly };
          scope.postMessage({ id: req.id, ok: true, result });
          break;
        }
        case 'apply': {
          if (!adapter) throw new Error('No document loaded');
          scope.postMessage({ id: req.id, ok: true, result: adapter.applyEdit(req.edit) });
          break;
        }
        case 'save': {
          if (!adapter) throw new Error('No document loaded');
          // An unchanged document returns the adapter's own buffer, which must stay usable: send a
          // copy, but hand it over (transfer) instead of cloning it a second time.
          const bytes = (await adapter.save()).slice();
          scope.postMessage({ id: req.id, ok: true, result: bytes }, [bytes.buffer]);
          break;
        }
      }
    } catch (err) {
      scope.postMessage({ id: req.id, ok: false, message: err instanceof Error ? err.message : String(err) });
    }
  }
}
