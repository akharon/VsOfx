declare module 'virtual:worker-bytes' {
  /** dist/worker.js, base64 encoded (inlined by esbuild.mjs). */
  const base64: string;
  export default base64;
}

declare module 'virtual:pdf-worker-bytes' {
  /** pdfjs-dist/build/pdf.worker.min.mjs, base64 encoded (inlined by esbuild.mjs). */
  const base64: string;
  export default base64;
}
