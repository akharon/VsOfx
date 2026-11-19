import type { CustomDocument, Uri } from 'vscode'; // types only, erased at build time
import { inverseEdit, type Edit } from '../shared/model';

/**
 * State of an open Office document: the original bytes and the list of
 * edits applied on top of them. The host never parses the bytes.
 * Has no runtime dependency on the `vscode` module so it can be unit tested.
 */
export class OfficeDocument implements CustomDocument {
  /** Edits currently applied, oldest first. */
  readonly edits: Edit[] = [];
  /** Number of edits that were applied when the document was last saved. */
  savedEditCount = 0;

  constructor(
    readonly uri: Uri,
    public originalBytes: Uint8Array,
  ) {}

  dispose(): void {
    // nothing to release; webviews are owned by the provider
  }

  /** Records an edit made by the user in a webview. */
  addEdit(edit: Edit): void {
    this.edits.push(edit);
  }

  /** Removes the most recent edit; returns the edit that undoes it. */
  undo(): Edit {
    const last = this.edits.pop();
    if (!last) throw new Error('Nothing to undo');
    return inverseEdit(last);
  }

  /** Re-applies an edit that was undone. */
  redo(edit: Edit): void {
    this.edits.push(edit);
  }

  markSaved(): void {
    this.savedEditCount = this.edits.length;
  }

  /** Replaces the base bytes (after revert) and drops all edits. */
  reset(bytes: Uint8Array): void {
    this.originalBytes = bytes;
    this.edits.length = 0;
    this.savedEditCount = 0;
  }

  /** Serialises state for hot-exit backups: [4-byte JSON length][JSON edits][original bytes]. */
  serialize(): Uint8Array {
    const json = new TextEncoder().encode(JSON.stringify({ edits: this.edits, savedEditCount: this.savedEditCount }));
    const out = new Uint8Array(4 + json.length + this.originalBytes.length);
    new DataView(out.buffer).setUint32(0, json.length, true);
    out.set(json, 4);
    out.set(this.originalBytes, 4 + json.length);
    return out;
  }

  /** Restores state written by `serialize`. */
  static deserialize(uri: OfficeDocument['uri'], data: Uint8Array): OfficeDocument {
    const len = new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(0, true);
    const meta = JSON.parse(new TextDecoder().decode(data.subarray(4, 4 + len))) as {
      edits: Edit[];
      savedEditCount: number;
    };
    const doc = new OfficeDocument(uri, data.slice(4 + len));
    doc.edits.push(...meta.edits);
    doc.savedEditCount = meta.savedEditCount;
    return doc;
  }
}
