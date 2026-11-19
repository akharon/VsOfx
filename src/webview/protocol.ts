import type { ApplyResult, DocModel, Edit, Format, SourcePart } from '../shared/model';

/** main thread -> worker */
export type WorkerRequest = { id: number } & (
  | { op: 'load'; format: Format; bytes: Uint8Array; edits: Edit[] }
  | { op: 'apply'; edit: Edit }
  | { op: 'save' }
  | { op: 'source' }
);

export interface LoadResult {
  model: DocModel;
  readonly: boolean;
}

/** worker -> main thread */
export type WorkerResponse =
  | { id: number; ok: true; result: LoadResult | ApplyResult | Uint8Array | SourcePart[] }
  | { id: number; ok: false; message: string };
