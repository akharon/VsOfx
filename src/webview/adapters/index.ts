import type { Format } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { DocxAdapter } from './docxAdapter';
import { OdsAdapter } from './odsAdapter';
import { OdtAdapter } from './odtAdapter';
import { PdfAdapter } from './pdfAdapter';
import { PptxAdapter } from './pptxAdapter';
import { XlsxAdapter } from './xlsxAdapter';
import { isZip } from './zip';

/** Creates the adapter for a format. Libraries are only touched when `load` runs. */
export function createAdapter(format: Format): FormatAdapter {
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
    case 'pdf':
      return new PdfAdapter();
  }
}

/** Throws a user-readable error if the bytes cannot be an Office package. */
export function assertOfficePackage(bytes: Uint8Array): void {
  if (bytes.length === 0) throw new Error('The file is empty.');
  if (!isZip(bytes)) throw new Error('This file is not a valid Office document (it is not a ZIP package).');
}
