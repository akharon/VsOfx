// Writes the synthetic sample files to test/fixtures/ (run: npm run fixtures; needs Node >= 22.18).
import { mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { makeDocx, makeOds, makeOdt, makePptx, makeXlsx, type DocxKind, type XlsxKind } from '../test/helpers/fixtures.ts';

const dir = fileURLToPath(new URL('../test/fixtures/', import.meta.url));
await mkdir(dir, { recursive: true });

const docx: DocxKind[] = ['simple', 'styles', 'tables', 'images', 'split-runs'];
for (const k of docx) await writeFile(dir + `${k}.docx`, makeDocx(k));
// large.xlsx (50k rows) is generated in the tests only, to keep the repository small
const xlsx: XlsxKind[] = ['simple', 'styles', 'formulas', 'multi-sheet'];
for (const k of xlsx) await writeFile(dir + `${k}.xlsx`, await makeXlsx(k));
await writeFile(dir + 'simple.pptx', makePptx());
await writeFile(dir + 'simple.odt', makeOdt());
await writeFile(dir + 'simple.ods', makeOds());
console.log(`fixtures written to ${dir}`);
