import type { ApplyResult, Edit } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { isZip } from './zip';

/** Batch edits are applied part by part (also while edits are replayed on load). */
export function withBatches(a: FormatAdapter): FormatAdapter {
  const apply = a.applyEdit.bind(a);
  a.applyEdit = (edit: Edit): ApplyResult => (edit.kind === 'batch' ? { batch: edit.edits.map((e) => a.applyEdit(e)) } : apply(edit));
  return a;
}

/** Throws a user-readable error if the bytes cannot be an Office package. */
export function assertOfficePackage(bytes: Uint8Array): void {
  if (bytes.length === 0) throw new Error('The file is empty.');
  if (!isZip(bytes)) throw new Error('This file is not a valid Office document (it is not a ZIP package).');
}

/** True if the bytes start like a PDF ("%PDF-" within the first KiB). */
export function isPdf(bytes: Uint8Array): boolean {
  const head = new TextDecoder('latin1').decode(bytes.subarray(0, 1024));
  return head.includes('%PDF-');
}

export function assertPdf(bytes: Uint8Array): void {
  if (bytes.length === 0) throw new Error('The file is empty.');
  if (!isPdf(bytes)) throw new Error('This file is not a valid PDF document.');
}
