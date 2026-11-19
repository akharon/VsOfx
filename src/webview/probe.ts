// TEST BUILD ONLY (esbuild.mjs --test-build; excluded from the .vsix).
// Tries to reach the network the way hostile third-party code would. Every attempt must be
// blocked by the Content-Security-Policy (doc/arch42.md §8.1); results go back through main.ts.
async function runProbe(): Promise<Record<string, string>> {
  const results: Record<string, string> = {};
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const target = 'https://example.com/vsofx-probe';

  try {
    await fetch(target, { mode: 'no-cors' });
    results['fetch'] = 'ALLOWED';
  } catch {
    results['fetch'] = 'blocked';
  }

  results['xhr'] = await new Promise<string>((resolve) => {
    try {
      const x = new XMLHttpRequest();
      x.onload = () => resolve('ALLOWED');
      x.onerror = () => resolve('blocked');
      x.open('GET', target);
      x.send();
    } catch {
      resolve('blocked');
    }
  });

  results['websocket'] = await new Promise<string>((resolve) => {
    try {
      const ws = new WebSocket('wss://example.com/vsofx-probe');
      ws.onopen = () => resolve('ALLOWED');
      ws.onerror = () => resolve('blocked');
      ws.onclose = () => resolve('blocked');
    } catch {
      resolve('blocked');
    }
  });

  results['image'] = await new Promise<string>((resolve) => {
    const img = new Image();
    img.onload = () => resolve('ALLOWED');
    img.onerror = () => resolve('blocked');
    img.src = 'https://example.com/vsofx-probe.png';
  });

  try {
    results['beacon'] = navigator.sendBeacon(target, 'x') ? 'ALLOWED' : 'blocked';
  } catch {
    results['beacon'] = 'blocked';
  }

  try {
    const f = document.createElement('iframe');
    f.src = target;
    document.body.appendChild(f);
    await wait(300);
    f.remove();
    results['iframe'] = 'requested'; // judged by the violation event (frame-src)
  } catch {
    results['iframe'] = 'blocked';
  }

  try {
    // eslint-disable-next-line no-eval
    results['eval'] = String(eval('1+1')) === '2' ? 'ALLOWED' : 'blocked';
  } catch {
    results['eval'] = 'blocked';
  }

  await wait(300); // let securitypolicyviolation events arrive
  return results;
}

window.addEventListener('vsofx-probe', () => {
  void runProbe().then((detail) => window.dispatchEvent(new CustomEvent('vsofx-probe-result', { detail })));
});
