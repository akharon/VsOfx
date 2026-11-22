declare module 'virtual:worker-source' {
  /** dist/worker-<family>.raw.js as a JavaScript string (inlined by esbuild.mjs). */
  const source: string;
  export default source;
}

declare module 'virtual:pdf-worker-source' {
  /** pdfjs-dist/legacy/build/pdf.worker.min.mjs as a JavaScript string (inlined by esbuild.mjs). */
  const source: string;
  export default source;
}
