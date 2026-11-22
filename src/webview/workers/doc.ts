// Worker of the document and presentation family: docx (mammoth), odt, pptx, odp/odg.
import type { Format } from '../../shared/model';
import type { FormatAdapter } from '../adapters/Adapter';
import { DocxAdapter } from '../adapters/docxAdapter';
import { OdpAdapter } from '../adapters/odpAdapter';
import { OdtAdapter } from '../adapters/odtAdapter';
import { PptxAdapter } from '../adapters/pptxAdapter';
import { runWorker } from './core';

runWorker((format: Format): FormatAdapter => {
  switch (format) {
    case 'docx':
      return new DocxAdapter();
    case 'odt':
      return new OdtAdapter();
    case 'pptx':
      return new PptxAdapter();
    case 'odp':
      return new OdpAdapter('presentation');
    case 'odg':
      return new OdpAdapter('drawing');
    default:
      throw new Error(`The document worker cannot open ${format} files.`);
  }
});
