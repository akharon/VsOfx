import type { ApplyResult, Edit, Format } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { CodeAdapter } from './codeAdapter';
import { CsvAdapter } from './csvAdapter';
import { DocxAdapter } from './docxAdapter';
import { OdpAdapter } from './odpAdapter';
import { OdsAdapter } from './odsAdapter';
import { OdtAdapter } from './odtAdapter';
import { PdfAdapter } from './pdfAdapter';
import { PptxAdapter } from './pptxAdapter';
import { XlsxAdapter } from './xlsxAdapter';
import { isZip } from './zip';

/** Creates the adapter for a format. Libraries are only touched when `load` runs. */
export function createAdapter(format: Format): FormatAdapter {
  return withBatches(adapterFor(format));
}

/** Batch edits are applied part by part (also while edits are replayed on load). */
function withBatches(a: FormatAdapter): FormatAdapter {
  const apply = a.applyEdit.bind(a);
  a.applyEdit = (edit: Edit): ApplyResult => (edit.kind === 'batch' ? { batch: edit.edits.map((e) => a.applyEdit(e)) } : apply(edit));
  return a;
}

function adapterFor(format: Format): FormatAdapter {
  switch (format) {
    case 'xlsx':
      return new XlsxAdapter();
    case 'docx':
      return new DocxAdapter();
    case 'pptx':
      return new PptxAdapter();
    case 'odt':
      return new OdtAdapter();
    case 'ods':
      return new OdsAdapter();
    case 'odp':
      return new OdpAdapter('presentation');
    case 'odg':
      return new OdpAdapter('drawing');
    case 'pdf':
      return new PdfAdapter();
    case 'csv':
      return new CsvAdapter();
    case 'markdown':
      return new CodeAdapter(format);
  }
}

/** Throws a user-readable error if the bytes cannot be an Office package. */
export function assertOfficePackage(bytes: Uint8Array): void {
  if (bytes.length === 0) throw new Error('The file is empty.');
  if (!isZip(bytes)) throw new Error('This file is not a valid Office document (it is not a ZIP package).');
}
