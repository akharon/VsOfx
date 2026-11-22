import type { Format } from '../shared/model';

/** Which format worker (dist/worker-<family>.js) opens a format. */
export type WorkerFamily = 'sheet' | 'doc' | 'pdf';

export function workerFamily(format: Format): WorkerFamily {
  switch (format) {
    case 'xlsx':
    case 'ods':
    case 'csv':
      return 'sheet';
    case 'pdf':
      return 'pdf';
    default:
      return 'doc';
  }
}
