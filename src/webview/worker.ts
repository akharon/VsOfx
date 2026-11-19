// Web Worker: all third-party parsing happens here (doc/arch42.md AD-7).
import { assertOfficePackage, createAdapter } from './adapters';
import { assertPdf } from './adapters/pdfAdapter';
import type { FormatAdapter } from './adapters/Adapter';
import { sourceParts } from './adapters/source';
import { isTextFormat, type Format } from '../shared/model';
import type { LoadResult, WorkerRequest, WorkerResponse } from './protocol';

let adapter: FormatAdapter | undefined;
let format: Format | undefined;

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
        adapter = createAdapter(req.format);
        format = req.format;
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
        // copied, not transferred: an unchanged document returns the adapter's own buffer
        const bytes = await adapter.save();
        scope.postMessage({ id: req.id, ok: true, result: bytes });
        break;
      }
      case 'source': {
        if (!adapter || !format) throw new Error('No document loaded');
        // the document as it would be saved now (unsaved edits included)
        scope.postMessage({ id: req.id, ok: true, result: sourceParts(format, await adapter.save()) });
        break;
      }
    }
  } catch (err) {
    scope.postMessage({ id: req.id, ok: false, message: err instanceof Error ? err.message : String(err) });
  }
}
