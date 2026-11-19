import { unzipSync, zipSync, strFromU8, strToU8, type Zippable } from 'fflate';

export type ZipEntries = Record<string, Uint8Array>;

const MAX_UNPACKED_BYTES = 512 * 1024 * 1024; // zip-bomb guard

/** True if the bytes start with a ZIP local file header ("PK\x03\x04"). */
export function isZip(b: Uint8Array): boolean {
  return b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;
}

/** Unzips all entries (insertion order = archive order). */
export function readZip(bytes: Uint8Array): ZipEntries {
  let total = 0;
  return unzipSync(bytes, {
    filter(file) {
      total += file.originalSize;
      if (total > MAX_UNPACKED_BYTES) throw new Error('Archive is too large when unpacked');
      return true;
    },
  });
}

export function text(entries: ZipEntries, name: string): string {
  const e = entries[name];
  if (!e) throw new Error(`Missing part "${name}" in the file`);
  return strFromU8(e);
}

export function optionalText(entries: ZipEntries, name: string): string | undefined {
  const e = entries[name];
  return e ? strFromU8(e) : undefined;
}

const ALREADY_COMPRESSED = /\.(png|jpe?g|gif|webp|mp3|mp4|m4a|zip|docx|xlsx|pptx)$/i;

/**
 * Writes a new archive: all entries of `entries`, with `replacements` swapped in.
 * Entry order is kept; with `mimetypeFirst` (ODF) the `mimetype` entry is written
 * first and stored uncompressed, as the ODF spec requires.
 */
export function writeZip(
  entries: ZipEntries,
  replacements: Record<string, Uint8Array | string>,
  opts: { mimetypeFirst?: boolean } = {},
): Uint8Array {
  const out: Zippable = {};
  const add = (name: string) => {
    const r = replacements[name];
    const data = r === undefined ? entries[name] : typeof r === 'string' ? strToU8(r) : r;
    if (!data) return;
    const stored = name === 'mimetype' || ALREADY_COMPRESSED.test(name);
    out[name] = [data, { level: stored ? 0 : 6 }];
  };
  if (opts.mimetypeFirst && entries['mimetype']) add('mimetype');
  for (const name of Object.keys(entries)) if (!(name in out)) add(name);
  for (const name of Object.keys(replacements)) if (!(name in out)) add(name);
  return zipSync(out);
}
