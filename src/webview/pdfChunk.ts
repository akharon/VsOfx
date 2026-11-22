// Entry of dist/pdf.js: pdf.js, the PDF view, the pdf.js worker source and the source of the PDF format
// worker (pdf-lib). Built as a classic script and loaded by main.ts only when a PDF is opened; it hands
// its exports over on globalThis.
import pdfWorkerSource from 'virtual:pdf-worker-source';
import formatWorkerSource from 'virtual:worker-source';
import { renderPdfView } from './views/pdfView';

export interface PdfChunk {
  renderPdfView: typeof renderPdfView;
  /** pdf.js worker (rendering). */
  workerSource: string;
  /** Format worker (pdf-lib: save, edits). */
  formatWorkerSource: string;
}

(globalThis as { vsofxPdf?: PdfChunk }).vsofxPdf = { renderPdfView, workerSource: pdfWorkerSource, formatWorkerSource };
