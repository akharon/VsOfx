import type { ApplyResult, DocModel, Edit } from '../../shared/model';

/** One adapter per file format; runs inside the Web Worker (doc/arch42.md §5). */
export interface FormatAdapter {
  /** True if the format cannot be edited (pptx). */
  readonly readonly: boolean;
  /** Parses the file, replays `edits` (restoring unsaved state) and returns the view model. */
  load(bytes: Uint8Array, edits: Edit[]): Promise<DocModel>;
  /** Applies one edit to the in-memory document. */
  applyEdit(edit: Edit): ApplyResult;
  /** Returns the new file bytes: original package plus all edits. */
  save(): Promise<Uint8Array>;
}
