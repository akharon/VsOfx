import { toByteArray } from './base64';

/**
 * File bytes in a message between the host and the webview. VS Code transfers typed arrays and
 * ArrayBuffers as they are (no JSON, no base64); other shapes are accepted too in case a VS Code
 * version serialises them differently: a Node Buffer as {type:'Buffer',data:[...]}, a number array, and
 * the base64 string older versions of this extension sent. Anything else throws.
 */
export function asBytes(v: unknown): Uint8Array {
  if (ArrayBuffer.isView(v)) return v instanceof Uint8Array ? v : new Uint8Array(v.buffer, v.byteOffset, v.byteLength);
  if (Object.prototype.toString.call(v) === '[object ArrayBuffer]') return new Uint8Array(v as ArrayBuffer);
  if (typeof v === 'string') return toByteArray(v);
  if (Array.isArray(v)) return Uint8Array.from(v as number[]);
  if (typeof v === 'object' && v !== null && (v as { type?: unknown }).type === 'Buffer' && Array.isArray((v as { data?: unknown }).data)) {
    return Uint8Array.from((v as { data: number[] }).data);
  }
  throw new Error('The message does not contain file bytes.');
}

/** Bytes that exactly fill their own ArrayBuffer, so that buffer can be transferred (copies if it is a view). */
export function ownedBytes(v: unknown): Uint8Array {
  const b = asBytes(v);
  return b.byteOffset === 0 && b.byteLength === b.buffer.byteLength && Object.prototype.toString.call(b.buffer) === '[object ArrayBuffer]' ? b : b.slice();
}
