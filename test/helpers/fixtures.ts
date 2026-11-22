// Builds small, valid Office files in memory (used by unit tests and by `npm run fixtures`).
// Real-Office/LibreOffice samples are still part of the manual checklist (doc/plan.md 6.1).
import ExcelJS from 'exceljs';
import { strToU8, zipSync, type Zippable } from 'fflate';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
// 1x1 transparent PNG
export const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

export function zip(files: Record<string, string | Uint8Array>, order?: string[]): Uint8Array {
  const z: Zippable = {};
  for (const name of order ?? Object.keys(files)) {
    const v = files[name]!;
    z[name] = typeof v === 'string' ? strToU8(v) : v;
  }
  return zipSync(z);
}

export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------------------------------------------------------------- docx

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

export type DocxKind = 'simple' | 'styles' | 'tables' | 'images' | 'split-runs';

const run = (t: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const para = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;

export function makeDocx(kind: DocxKind = 'simple'): Uint8Array {
  let body = '';
  switch (kind) {
    case 'simple':
      body = para(run('First paragraph.')) + para(run('Second paragraph with <angle> & ampersand.')) + para('') + para(run('Last paragraph.'));
      break;
    case 'styles':
      body =
        para(run('Main Title'), '<w:pStyle w:val="Heading1"/>') +
        para(run('Intro text.')) +
        para(run('Bullet one'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        para(run('Bullet two'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        para(run('Subtitle'), '<w:pStyle w:val="Heading2"/>') +
        para(run('Closing text.'));
      break;
    case 'tables':
      body =
        para(run('Before table')) +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>' +
        '<w:tr><w:tc>' + para(run('A1')) + '</w:tc><w:tc>' + para(run('B1')) + '</w:tc></w:tr>' +
        '<w:tr><w:tc>' + para(run('A2')) + '</w:tc><w:tc>' + para(run('B2')) + '</w:tc></w:tr>' +
        '</w:tbl>' +
        para(run('After table'));
      break;
    case 'images':
      body =
        para(run('Text before image')) +
        para(
          '<w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/>' +
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
            '<pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
            '<pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>',
        ) +
        para(run('Text after image'));
      break;
    case 'split-runs':
      body =
        para(run('Hello ', '<w:b/>') + run('wor', '<w:i/>') + run('ld', '')) +
        para(run('Plain second paragraph.'));
      break;
  }
  const files: Record<string, string | Uint8Array> = {
    '[Content_Types].xml':
      XML +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
    '_rels/.rels':
      XML +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels':
      XML +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
    'word/styles.xml':
      XML +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>' +
      '</w:styles>',
    'word/numbering.xml':
      XML +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    'word/media/image1.png': PNG,
    'word/document.xml': `${XML}<w:document ${W_NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
  };
  return zip(files);
}

// ---------------------------------------------------------------- xlsx

export type XlsxKind = 'simple' | 'styles' | 'formulas' | 'multi-sheet' | 'large';

export async function makeXlsx(kind: XlsxKind = 'simple'): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  switch (kind) {
    case 'simple':
      ws.addRow(['Name', 'Qty', 'Note']);
      ws.addRow(['Apples', 3, 'fresh']);
      ws.addRow(['Pears', 5, '']);
      break;
    case 'styles': {
      ws.addRow(['Header', 'Amount', 'Date']);
      ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
      ws.addRow(['Total', 1234.5, new Date(Date.UTC(2024, 0, 15))]);
      ws.getCell('B2').numFmt = '#,##0.00';
      ws.getCell('C2').numFmt = 'yyyy-mm-dd';
      ws.getCell('A2').font = { italic: true };
      ws.getCell('B2').alignment = { horizontal: 'center' };
      ws.getColumn(1).width = 24;
      ws.mergeCells('A4:C4');
      ws.getCell('A4').value = 'Merged title';
      break;
    }
    case 'formulas':
      ws.addRow([1, 2, { formula: 'A1+B1', result: 3 }]);
      ws.addRow([10, 20, { formula: 'SUM(A2:B2)', result: 30 }]);
      break;
    case 'multi-sheet':
      ws.addRow(['one']);
      wb.addWorksheet('Second').addRow(['two']);
      wb.addWorksheet('Third').addRow(['three']);
      break;
    case 'large':
      for (let i = 1; i <= 50_000; i++) ws.addRow([i, `row ${i}`, i * 1.5]);
      break;
  }
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

// ---------------------------------------------------------------- pptx

export function makePptx(): Uint8Array {
  const P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const sp = (type: string, lines: string[]) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr>${type ? `<p:ph type="${type}"/>` : '<p:ph/>'}</p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/>${lines.map((l) => `<a:p><a:r><a:t>${esc(l)}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;
  const slide = (title: string, body: string[], pic = false) =>
    `${XML}<p:sld ${P}><p:cSld><p:spTree>${sp('title', [title])}${sp('', body)}` +
    (pic ? '<p:pic><p:nvPicPr><p:cNvPr id="5" name="pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr/></p:pic>' : '') +
    `</p:spTree></p:cSld></p:sld>`;
  const rels = (inner: string) => `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return zip({
    '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
    '_rels/.rels': rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/>`),
    // slide ids deliberately in a different order than the file names
    'ppt/presentation.xml': `${XML}<p:presentation ${P}><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(`<Relationship Id="rId1" Type="${R}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${R}/slide" Target="slides/slide2.xml"/>`),
    'ppt/slides/slide1.xml': slide('Second in order', ['Detail line']),
    'ppt/slides/slide2.xml': slide('Opening slide', ['Point A', 'Point B'], true),
    'ppt/slides/_rels/slide2.xml.rels': rels(`<Relationship Id="rId5" Type="${R}/image" Target="../media/image1.png"/><Relationship Id="rId6" Type="${R}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>`),
    'ppt/notesSlides/notesSlide1.xml': `${XML}<p:notes ${P}><p:cSld><p:spTree>${sp('body', ['Remember to smile'])}</p:spTree></p:cSld></p:notes>`,
    'ppt/media/image1.png': PNG,
  });
}

// ---------------------------------------------------------------- odf

const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

function odf(mime: string, content: string): Uint8Array {
  return zip(
    {
      mimetype: mime,
      'content.xml': `${XML}<office:document-content ${ODF_NS} office:version="1.3">${content}</office:document-content>`,
      'META-INF/manifest.xml': `${XML}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mime}"/></manifest:manifest>`,
    },
    // deliberately NOT mimetype first: the adapter must fix the order on save
    ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  );
}

export function makeOdt(): Uint8Array {
  return odf(
    'application/vnd.oasis.opendocument.text',
    '<office:automatic-styles><style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles>' +
      '<office:body><office:text>' +
      '<text:h text:outline-level="1">Report title</text:h>' +
      '<text:p>Plain paragraph with <text:span text:style-name="T1">bold part</text:span> and a<text:s text:c="3"/>gap.</text:p>' +
      '<text:list><text:list-item><text:p>List item</text:p></text:list-item></text:list>' +
      '<table:table table:name="T"><table:table-row><table:table-cell><text:p>Cell text</text:p></table:table-cell></table:table-row></table:table>' +
      '<text:p>Last one</text:p>' +
      '</office:text></office:body>',
  );
}

export function makeOds(): Uint8Array {
  return odf(
    'application/vnd.oasis.opendocument.spreadsheet',
    '<office:body><office:spreadsheet>' +
      '<table:table table:name="Data">' +
      '<table:table-column table:number-columns-repeated="3"/>' +
      '<table:table-row>' +
      '<table:table-cell office:value-type="string"><text:p>Name</text:p></table:table-cell>' +
      '<table:table-cell office:value-type="string"><text:p>Value</text:p></table:table-cell></table:table-row>' +
      // a repeated row holding a repeated cell: tests splitting
      '<table:table-row table:number-rows-repeated="3">' +
      '<table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="2"><text:p>7</text:p></table:table-cell></table:table-row>' +
      '<table:table-row table:number-rows-repeated="1048000"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>' +
      '</table:table>' +
      '<table:table table:name="Second"><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>' +
      '</office:spreadsheet></office:body>',
  );
}

const PRES_NS = 'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"';

export function makeOdp(kind: 'presentation' | 'drawing' = 'presentation'): Uint8Array {
  const body = kind === 'presentation' ? 'presentation' : 'drawing';
  return odf(
    kind === 'presentation' ? 'application/vnd.oasis.opendocument.presentation' : 'application/vnd.oasis.opendocument.graphics',
    `<office:body ${PRES_NS}><office:${body}>` +
      '<draw:page draw:name="page1">' +
      '<draw:frame presentation:class="title"><draw:text-box><text:p>Hello ODP</text:p></draw:text-box></draw:frame>' +
      '<draw:frame presentation:class="outline"><draw:text-box><text:list><text:list-item><text:p>Point one</text:p></text:list-item></text:list></draw:text-box></draw:frame>' +
      '<draw:custom-shape><text:p>Shape text</text:p></draw:custom-shape>' +
      '<presentation:notes><draw:frame><draw:text-box><text:p>Speaker note</text:p></draw:text-box></draw:frame></presentation:notes>' +
      '</draw:page>' +
      `</office:${body}></office:body>`,
  );
}

export function makeOdsFormulas(): Uint8Array {
  return odf(
    'application/vnd.oasis.opendocument.spreadsheet',
    '<office:body><office:spreadsheet><table:table table:name="S">' +
      '<table:table-row>' +
      '<table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>' +
      '<table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>' +
      '<table:table-cell table:formula="of:=SUM([.A1:.B1])" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>' +
      '</table:table-row></table:table></office:spreadsheet></office:body>',
  );
}

/** Item / Qty table (b 3, a 10, c 2) with =SUM(B2:B4) in B6 and a styled cell, repeated trailing rows like LibreOffice writes. */
export function makeOdsData(): Uint8Array {
  const s = (t: string) => `<table:table-cell office:value-type="string"><text:p>${t}</text:p></table:table-cell>`;
  const n = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  return odf(
    'application/vnd.oasis.opendocument.spreadsheet',
    '<office:body><office:spreadsheet><table:table table:name="S">' +
      '<table:table-column table:number-columns-repeated="2"/>' +
      `<table:table-row>${s('Item')}${s('Qty')}</table:table-row>` +
      `<table:table-row>${s('b')}${n(3)}</table:table-row>` +
      `<table:table-row>${s('a')}${n(10)}</table:table-row>` +
      `<table:table-row>${s('c')}${n(2)}</table:table-row>` +
      '<table:table-row><table:table-cell table:number-columns-repeated="2"/></table:table-row>' +
      '<table:table-row><table:table-cell/><table:table-cell table:formula="of:=SUM([.B2:.B4])" office:value-type="float" office:value="15"><text:p>15</text:p></table:table-cell></table:table-row>' +
      '<table:table-row table:number-rows-repeated="1000"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>' +
      '</table:table></office:spreadsheet></office:body>',
  );
}
