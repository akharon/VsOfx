import type { Format } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { assertOfficePackage, withBatches } from './common';
import { CsvAdapter } from './csvAdapter';
import { DocxAdapter } from './docxAdapter';
import { OdpAdapter } from './odpAdapter';
import { OdsAdapter } from './odsAdapter';
import { OdtAdapter } from './odtAdapter';
import { PdfAdapter } from './pdfAdapter';
import { PptxAdapter } from './pptxAdapter';
import { XlsxAdapter } from './xlsxAdapter';

export { assertOfficePackage };

/**
 * Creates the adapter for a format (all families; the shipped workers import only their own, see
 * src/webview/workers/). Libraries are only touched when `load` runs.
 */
export function createAdapter(format: Format): FormatAdapter {
  return withBatches(adapterFor(format));
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
  }
}
