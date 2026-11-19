export interface WebviewHtmlOptions {
  cspSource: string;
  scriptUri: string;
  styleUri: string;
  /** Extra script (test builds only). */
  probeUri?: string;
  nonce: string;
}

/** Content-Security-Policy for the editor webview (doc/arch42.md §8.1). No network, no eval. */
export function buildCsp(cspSource: string, nonce: string): string {
  return [
    `default-src 'none'`,
    `script-src 'nonce-${nonce}'`,
    `style-src ${cspSource} 'unsafe-inline'`,
    `img-src data: blob:`,
    `font-src ${cspSource}`,
    `worker-src blob: ${cspSource}`,
    `connect-src 'none'`,
    `form-action 'none'`,
    `frame-src 'none'`,
  ].join('; ');
}

export function makeNonce(): string {
  const bytes = new Uint8Array(16);
  globalThis.crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

export function buildWebviewHtml(o: WebviewHtmlOptions): string {
  const probe = o.probeUri ? `\n<script nonce="${o.nonce}" src="${o.probeUri}"></script>` : '';
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${buildCsp(o.cspSource, o.nonce)}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${o.styleUri}">
<title>VsOfx</title>
</head>
<body>
<div id="app" aria-live="polite"><div class="loading">Loading…</div></div>
<script nonce="${o.nonce}" src="${o.scriptUri}"></script>${probe}
</body>
</html>`;
}
