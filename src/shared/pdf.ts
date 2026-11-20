// PDF edit model, shared by host (validation), worker (save) and view.
//
// Every PdfOp is an absolute assignment ("field X has value V", "widget W is removed"), so applying
// ops in any order any number of times gives the same state, and an undo is just the op with
// `old`/`value` swapped. The saved file is always "original bytes + current state".

/** Rectangle in PDF user space: [x1, y1, x2, y2]. */
export type Rect = [number, number, number, number];

export type FieldKind = 'text' | 'checkbox' | 'dropdown';

/** A form field created by the user (original fields come from the file). */
export interface FieldSpec {
  id: string; // unique, "n1", "n2" ...
  kind: FieldKind;
  name: string;
  page: number; // 0-based, original page index
  rect: Rect;
  options?: string[]; // dropdown
  multiline?: boolean; // text
  size?: number; // text: font size (unset = automatic)
  bold?: boolean;
  italic?: boolean;
  color?: string; // "#rrggbb"
}

const optBool = (v: unknown) => v === undefined || typeof v === 'boolean';
const optColor = (v: unknown) => v === undefined || (typeof v === 'string' && /^#[0-9a-f]{6}$/i.test(v));

/** Free text placed on a page (printed into the page content on save). */
export interface TextSpec {
  id: string; // "t1", "t2" ...
  page: number;
  rect: Rect; // x1,y1 = lower-left corner
  text: string;
  size: number;
  bold?: boolean;
  italic?: boolean;
  color?: string; // "#rrggbb"
}

/** An image placed on a page (drawn into the page content on save). */
export interface ImageSpec {
  id: string; // "i1", "i2" ...
  page: number;
  rect: Rect;
  mime: 'image/png' | 'image/jpeg';
  data: string; // base64
}

export type FieldValue = string | boolean;

export type PdfOp =
  | { op: 'setField'; name: string; value: FieldValue; old: FieldValue }
  | { op: 'addField'; spec: FieldSpec }
  | { op: 'dropField'; spec: FieldSpec } // inverse of addField
  | { op: 'removeWidget'; id: string; name: string } // delete an original form field widget
  | { op: 'restoreWidget'; id: string; name: string }
  | { op: 'setRect'; id: string; rect: Rect; old: Rect }
  | { op: 'addText'; spec: TextSpec }
  | { op: 'dropText'; spec: TextSpec }
  | { op: 'addImage'; spec: ImageSpec }
  | { op: 'dropImage'; spec: ImageSpec }
  | { op: 'setText'; id: string; text: string; old: string }
  | { op: 'setRotation'; page: number; deg: number; old: number } // extra rotation, multiple of 90
  | { op: 'deletePage'; page: number }
  | { op: 'restorePage'; page: number };

export interface PdfEdit {
  kind: 'pdf';
  op: PdfOp;
}

export function inverseOp(o: PdfOp): PdfOp {
  switch (o.op) {
    case 'setField':
      return { ...o, value: o.old, old: o.value };
    case 'addField':
      return { op: 'dropField', spec: o.spec };
    case 'dropField':
      return { op: 'addField', spec: o.spec };
    case 'removeWidget':
      return { ...o, op: 'restoreWidget' };
    case 'restoreWidget':
      return { ...o, op: 'removeWidget' };
    case 'setRect':
      return { ...o, rect: o.old, old: o.rect };
    case 'addText':
      return { op: 'dropText', spec: o.spec };
    case 'dropText':
      return { op: 'addText', spec: o.spec };
    case 'addImage':
      return { op: 'dropImage', spec: o.spec };
    case 'dropImage':
      return { op: 'addImage', spec: o.spec };
    case 'setText':
      return { ...o, text: o.old, old: o.text };
    case 'setRotation':
      return { ...o, deg: o.old, old: o.deg };
    case 'deletePage':
      return { op: 'restorePage', page: o.page };
    case 'restorePage':
      return { op: 'deletePage', page: o.page };
  }
}

// ---- validation (the host does not trust webview messages) ----

const int = (v: unknown): v is number => typeof v === 'number' && Number.isInteger(v) && v >= 0;
const num = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
const str = (v: unknown): v is string => typeof v === 'string' && v.length <= 100_000;
const rect = (v: unknown): v is Rect => Array.isArray(v) && v.length === 4 && v.every(num);
const value = (v: unknown): v is FieldValue => typeof v === 'boolean' || str(v);

function isFieldSpec(v: unknown): v is FieldSpec {
  if (typeof v !== 'object' || v === null) return false;
  const s = v as Record<string, unknown>;
  return (
    str(s['id']) &&
    (s['kind'] === 'text' || s['kind'] === 'checkbox' || s['kind'] === 'dropdown') &&
    str(s['name']) &&
    int(s['page']) &&
    rect(s['rect']) &&
    (s['options'] === undefined || (Array.isArray(s['options']) && s['options'].every(str))) &&
    (s['multiline'] === undefined || typeof s['multiline'] === 'boolean') &&
    (s['size'] === undefined || (num(s['size']) && (s['size'] as number) > 0 && (s['size'] as number) <= 200)) &&
    optBool(s['bold']) &&
    optBool(s['italic']) &&
    optColor(s['color'])
  );
}

function isTextSpec(v: unknown): v is TextSpec {
  if (typeof v !== 'object' || v === null) return false;
  const s = v as Record<string, unknown>;
  return str(s['id']) && int(s['page']) && rect(s['rect']) && str(s['text']) && num(s['size']) && (s['size'] as number) > 0 &&
    (s['bold'] === undefined || typeof s['bold'] === 'boolean') &&
    (s['italic'] === undefined || typeof s['italic'] === 'boolean') &&
    (s['color'] === undefined || (typeof s['color'] === 'string' && /^#[0-9a-f]{6}$/i.test(s['color'])));
}

/**
 * Largest image accepted, as base64 characters (~9 MB of image data). Each copy of an addImage op
 * (undo stack, mirrored editors, hot-exit backup) carries the whole image, so the view reduces large
 * images before inserting them.
 */
export const MAX_IMAGE_BASE64 = 12_000_000;

function isImageSpec(v: unknown): v is ImageSpec {
  if (typeof v !== 'object' || v === null) return false;
  const s = v as Record<string, unknown>;
  const d = s['data'];
  return (
    str(s['id']) &&
    int(s['page']) &&
    rect(s['rect']) &&
    (s['mime'] === 'image/png' || s['mime'] === 'image/jpeg') &&
    typeof d === 'string' &&
    d.length <= MAX_IMAGE_BASE64 &&
    /^[A-Za-z0-9+/=]*$/.test(d)
  );
}

export function isPdfOp(x: unknown): x is PdfOp {
  if (typeof x !== 'object' || x === null) return false;
  const o = x as Record<string, unknown>;
  switch (o['op']) {
    case 'setField':
      return str(o['name']) && value(o['value']) && value(o['old']);
    case 'addField':
    case 'dropField':
      return isFieldSpec(o['spec']);
    case 'removeWidget':
    case 'restoreWidget':
      return str(o['id']) && str(o['name']);
    case 'setRect':
      return str(o['id']) && rect(o['rect']) && rect(o['old']);
    case 'addText':
    case 'dropText':
      return isTextSpec(o['spec']);
    case 'addImage':
    case 'dropImage':
      return isImageSpec(o['spec']);
    case 'setText':
      return str(o['id']) && str(o['text']) && str(o['old']);
    case 'setRotation':
      return int(o['page']) && num(o['deg']) && num(o['old']);
    case 'deletePage':
    case 'restorePage':
      return int(o['page']);
    default:
      return false;
  }
}

export function isPdfEdit(x: unknown): x is PdfEdit {
  return typeof x === 'object' && x !== null && (x as { kind?: unknown }).kind === 'pdf' && isPdfOp((x as { op?: unknown }).op);
}

// ---- state ----

export interface PdfState {
  values: Map<string, FieldValue>;
  added: Map<string, FieldSpec>;
  removed: Set<string>; // original widget ids
  rects: Map<string, Rect>; // moved/resized widgets and text boxes, by id
  texts: Map<string, TextSpec>;
  images: Map<string, ImageSpec>;
  rotations: Map<number, number>;
  deletedPages: Set<number>;
}

export function emptyPdfState(): PdfState {
  return {
    values: new Map(),
    added: new Map(),
    removed: new Set(),
    rects: new Map(),
    texts: new Map(),
    images: new Map(),
    rotations: new Map(),
    deletedPages: new Set(),
  };
}

export function applyOp(s: PdfState, o: PdfOp): void {
  switch (o.op) {
    case 'setField':
      s.values.set(o.name, o.value);
      break;
    case 'addField':
      s.added.set(o.spec.id, o.spec);
      break;
    case 'dropField':
      s.added.delete(o.spec.id);
      break;
    case 'removeWidget':
      s.removed.add(o.id);
      break;
    case 'restoreWidget':
      s.removed.delete(o.id);
      break;
    case 'setRect':
      s.rects.set(o.id, o.rect);
      break;
    case 'addText':
      s.texts.set(o.spec.id, o.spec);
      break;
    case 'dropText':
      s.texts.delete(o.spec.id);
      break;
    case 'addImage':
      s.images.set(o.spec.id, o.spec);
      break;
    case 'dropImage':
      s.images.delete(o.spec.id);
      break;
    case 'setText': {
      const t = s.texts.get(o.id);
      if (t) s.texts.set(o.id, { ...t, text: o.text });
      break;
    }
    case 'setRotation':
      if (o.deg % 360 === 0) s.rotations.delete(o.page);
      else s.rotations.set(o.page, ((o.deg % 360) + 360) % 360);
      break;
    case 'deletePage':
      s.deletedPages.add(o.page);
      break;
    case 'restorePage':
      s.deletedPages.delete(o.page);
      break;
  }
}

export function reduceOps(ops: Iterable<PdfOp>): PdfState {
  const s = emptyPdfState();
  for (const o of ops) applyOp(s, o);
  return s;
}

/** Serialisable snapshot of the state (sent worker -> view). */
export interface PdfStateJson {
  values: [string, FieldValue][];
  added: FieldSpec[];
  removed: string[];
  rects: [string, Rect][];
  texts: TextSpec[];
  images?: ImageSpec[];
  rotations: [number, number][];
  deletedPages: number[];
}

export function stateToJson(s: PdfState): PdfStateJson {
  return {
    values: [...s.values],
    added: [...s.added.values()],
    removed: [...s.removed],
    rects: [...s.rects],
    texts: [...s.texts.values()],
    images: [...s.images.values()],
    rotations: [...s.rotations],
    deletedPages: [...s.deletedPages],
  };
}

export function stateFromJson(j: PdfStateJson): PdfState {
  return {
    values: new Map(j.values),
    added: new Map(j.added.map((a) => [a.id, a])),
    removed: new Set(j.removed),
    rects: new Map(j.rects),
    texts: new Map(j.texts.map((t) => [t.id, t])),
    images: new Map((j.images ?? []).map((i) => [i.id, i])),
    rotations: new Map(j.rotations),
    deletedPages: new Set(j.deletedPages),
  };
}

export function pdfEditLabel(o: PdfOp): string {
  switch (o.op) {
    case 'setField':
      return 'Fill form field';
    case 'addField':
    case 'dropField':
      return o.op === 'addField' ? 'Add form field' : 'Remove form field';
    case 'removeWidget':
      return 'Remove form field';
    case 'restoreWidget':
      return 'Restore form field';
    case 'setRect':
      return 'Move or resize';
    case 'addText':
      return 'Add text';
    case 'dropText':
      return 'Remove text';
    case 'addImage':
      return 'Add image';
    case 'dropImage':
      return 'Remove image';
    case 'setText':
      return 'Edit text';
    case 'setRotation':
      return 'Rotate page';
    case 'deletePage':
      return 'Delete page';
    case 'restorePage':
      return 'Restore page';
  }
}
