// "VsOfx: New Document": the smallest valid empty files, generated here (no template assets).
// The host bundle has no third-party code, so the archive is written by hand (stored, no compression).
export type NewDocKind = 'docx' | 'xlsx' | 'odt' | 'ods';

export const NEW_DOCS: readonly { kind: NewDocKind; label: string; detail: string }[] = [
  { kind: 'docx', label: 'Word document', detail: '.docx' },
  { kind: 'xlsx', label: 'Spreadsheet', detail: '.xlsx' },
  { kind: 'odt', label: 'Writer document', detail: '.odt' },
  { kind: 'ods', label: 'Calc spreadsheet', detail: '.ods' },
];

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(b: Uint8Array): number {
  let c = 0xffffffff;
  for (const x of b) c = CRC_TABLE[(c ^ x) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** A ZIP archive with the entries in the given order, stored uncompressed. */
export function storeZip(files: [name: string, content: string][]): Uint8Array {
  const enc = new TextEncoder();
  const entries = files.map(([name, content]) => {
    const data = enc.encode(content);
    return { name: enc.encode(name), data, crc: crc32(data), offset: 0 };
  });
  const size = entries.reduce((n, e) => n + 30 + e.name.length + e.data.length + 46 + e.name.length, 22);
  const out = new Uint8Array(size);
  const v = new DataView(out.buffer);
  let p = 0;
  for (const e of entries) {
    e.offset = p;
    v.setUint32(p, 0x04034b50, true);
    v.setUint16(p + 4, 20, true);
    v.setUint16(p + 6, 0x0800, true); // UTF-8 names
    v.setUint16(p + 12, 33, true); // 1980-01-01
    v.setUint32(p + 14, e.crc, true);
    v.setUint32(p + 18, e.data.length, true);
    v.setUint32(p + 22, e.data.length, true);
    v.setUint16(p + 26, e.name.length, true);
    out.set(e.name, p + 30);
    out.set(e.data, p + 30 + e.name.length);
    p += 30 + e.name.length + e.data.length;
  }
  const dir = p;
  for (const e of entries) {
    v.setUint32(p, 0x02014b50, true);
    v.setUint16(p + 4, 20, true);
    v.setUint16(p + 6, 20, true);
    v.setUint16(p + 8, 0x0800, true);
    v.setUint16(p + 14, 33, true);
    v.setUint32(p + 16, e.crc, true);
    v.setUint32(p + 20, e.data.length, true);
    v.setUint32(p + 24, e.data.length, true);
    v.setUint16(p + 28, e.name.length, true);
    v.setUint32(p + 42, e.offset, true);
    out.set(e.name, p + 46);
    p += 46 + e.name.length;
  }
  v.setUint32(p, 0x06054b50, true);
  v.setUint16(p + 8, entries.length, true);
  v.setUint16(p + 10, entries.length, true);
  v.setUint32(p + 12, p - dir, true);
  v.setUint32(p + 16, dir, true);
  return out;
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const OOXML = 'application/vnd.openxmlformats-officedocument';

const rels = (target: string, extra = '') =>
  `${XML}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="${target}"/>${extra}</Relationships>`;

function docx(): [string, string][] {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return [
    ['[Content_Types].xml', `${XML}<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OOXML}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OOXML}.wordprocessingml.styles+xml"/></Types>`],
    ['_rels/.rels', rels('word/document.xml')],
    ['word/document.xml', `${XML}<w:document xmlns:w="${W}"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`],
    ['word/_rels/document.xml.rels', `${XML}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ['word/styles.xml', `${XML}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`],
  ];
}

function xlsx(): [string, string][] {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  return [
    ['[Content_Types].xml', `${XML}<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${OOXML}.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${OOXML}.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', rels('xl/workbook.xml')],
    ['xl/workbook.xml', `${XML}<workbook xmlns="${S}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `${XML}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`],
    ['xl/worksheets/sheet1.xml', `${XML}<worksheet xmlns="${S}"><sheetData/></worksheet>`],
  ];
}

function odf(mime: string, body: string): [string, string][] {
  const ns =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
  return [
    ['mimetype', mime], // first, as ODF requires
    ['content.xml', `${XML}<office:document-content ${ns} office:version="1.2"><office:body>${body}</office:body></office:document-content>`],
    ['META-INF/manifest.xml', `${XML}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${mime}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`],
  ];
}

/** The bytes of a new, empty document. */
export function newDocumentBytes(kind: NewDocKind): Uint8Array {
  switch (kind) {
    case 'docx':
      return storeZip(docx());
    case 'xlsx':
      return storeZip(xlsx());
    case 'odt':
      return storeZip(odf('application/vnd.oasis.opendocument.text', '<office:text><text:p/></office:text>'));
    case 'ods':
      return storeZip(
        odf(
          'application/vnd.oasis.opendocument.spreadsheet',
          '<office:spreadsheet><table:table table:name="Sheet1"><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet>',
        ),
      );
  }
}
