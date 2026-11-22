/**
 * Loading skeleton: a calm placeholder (toolbar shell plus a grid or a page) that the host puts into the
 * page HTML so it shows before any script has run, and that main.ts shows again while a document loads.
 * Static markup only (no data from files); the styles are in src/webview/style.css.
 */
export type SkeletonKind = 'grid' | 'page';

export function skeletonKind(format: string | undefined): SkeletonKind {
  return format === 'xlsx' || format === 'ods' || format === 'csv' ? 'grid' : 'page';
}

const LINES = [92, 100, 84, 97, 60, 0, 100, 88, 95, 72, 0, 99, 90, 46];

export function skeletonHtml(kind: SkeletonKind): string {
  const toolbar = `<div class="sk-toolbar" aria-hidden="true">${'<i></i>'.repeat(9)}</div>`;
  const body =
    kind === 'grid'
      ? '<div class="sk-grid" aria-hidden="true"></div>'
      : `<div class="sk-desk" aria-hidden="true"><div class="sk-page">${LINES.map((w) => (w ? `<i style="width:${w}%"></i>` : '<b></b>')).join('')}</div></div>`;
  return `<div class="skeleton sk-${kind}" role="status"><span class="sk-label">Loading…</span>${toolbar}${body}</div>`;
}
