import type { ApplyResult, Edit } from '../../shared/model';

export interface View {
  /** Called after an edit was applied in the worker (user edit or host undo/redo). */
  applyEdit(edit: Edit, result: ApplyResult): void;
  /** Commits input that is still being typed (called before save). */
  flush?(): void;
  dispose(): void;
}

export type EditHandler = (edit: Edit) => void;
