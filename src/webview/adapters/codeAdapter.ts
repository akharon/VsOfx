import { applyTextEdit, type ApplyResult, type CodeLang, type DocModel, type Edit } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { decodeText, encodeText, type TextEncodingInfo } from './textCodec';

/**
 * JSON, YAML, XML and Markdown: the document is its text. Edits are splices, so the
 * user's formatting, comments and key order are never touched by a save.
 */
export class CodeAdapter implements FormatAdapter {
  readonly readonly = false;
  private original: Uint8Array = new Uint8Array(0);
  private text = '';
  private loadedText = '';
  private info!: TextEncodingInfo;

  constructor(private readonly lang: CodeLang) {}

  async load(bytes: Uint8Array, edits: Edit[]): Promise<DocModel> {
    const { text, info } = decodeText(bytes);
    this.original = bytes;
    this.info = info;
    this.loadedText = text;
    this.text = text;
    for (const e of edits) this.applyEdit(e);
    return { type: 'code', lang: this.lang, text: this.text };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'text') throw new Error('Not a text edit');
    this.text = applyTextEdit(this.text, edit);
    return {};
  }

  async save(): Promise<Uint8Array> {
    if (this.text === this.loadedText) return this.original;
    return encodeText(this.text, this.info);
  }
}
