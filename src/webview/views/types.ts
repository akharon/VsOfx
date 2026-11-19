import type { ApplyResult, Edit } from '../../shared/model';

export interface View {
  /** Called after an edit was applied in the worker (user edit or host undo/redo). */
  applyEdit(edit: Edit, result: ApplyResult): void;
  dispose(): void;
}

export type EditHandler = (edit: Edit) => void;
