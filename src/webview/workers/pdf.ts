// Worker of the PDF family (pdf-lib).
import type { Format } from '../../shared/model';
import type { FormatAdapter } from '../adapters/Adapter';
import { PdfAdapter } from '../adapters/pdfAdapter';
import { runWorker } from './core';

runWorker((format: Format): FormatAdapter => {
  if (format !== 'pdf') throw new Error(`The PDF worker cannot open ${format} files.`);
  return new PdfAdapter();
});
