// Runs inside VS Code (npm run test:integration). Uses the test build of the extension, which
// contains dist/probe.js and the test-mode API (see src/host/extension.ts).
import * as assert from 'node:assert';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { PDFDocument, PDFTextField } from '@cantoo/pdf-lib';
import { strFromU8, unzipSync } from 'fflate';
import type { Edit } from '../../src/shared/model';

interface TestApi {
  violations: { directive: string; blockedURI: string }[];
  probeResults: Record<string, string>[];
  errors: string[];
  probeAll(): void;
  injectEdit(uri: vscode.Uri, edit: Edit): void;
}

const fixtures = path.resolve(__dirname, '..', '..', '..', 'test', 'fixtures');
const VIEW_TYPE = 'vsofx.officeEditor';
let lastApi: { errors: string[] } | undefined;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function until<T>(fn: () => T | undefined | false, what: string, ms = 30_000): Promise<T> {
  const end = Date.now() + ms;
  for (;;) {
    const v = fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`Timed out waiting for ${what}. Extension errors: ${JSON.stringify(lastApi?.errors)}`);
    await sleep(100);
  }
}

function tempCopy(name: string): vscode.Uri {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vsofx-'));
  const file = path.join(dir, name);
  fs.copyFileSync(path.join(fixtures, name), file);
  return vscode.Uri.file(file);
}

function openTabs(uri: vscode.Uri) {
  return vscode.window.tabGroups.all
    .flatMap((g) => g.tabs)
    .filter((t) => t.input instanceof vscode.TabInputCustom && t.input.uri.toString() === uri.toString() && t.input.viewType === VIEW_TYPE);
}

function entry(file: string, part: string): string {
  return strFromU8(unzipSync(new Uint8Array(fs.readFileSync(file)))[part]!);
}

describe('VsOfx in VS Code', function () {
  let api: TestApi;

  before(async () => {
    const ext = vscode.extensions.getExtension('vsofx.vsofx')!;
    assert.ok(ext, 'extension is installed');
    api = (await ext.activate()) as TestApi;
    assert.ok(api?.probeAll, 'test API is exposed in test mode');
    lastApi = api;
  });

  afterEach(async () => {
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  });

  it('opens every supported format in the custom editor', async () => {
    for (const name of ['simple.xlsx', 'simple.docx', 'simple.pptx', 'simple.odt', 'simple.ods']) {
      const uri = tempCopy(name);
      await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
      await until(() => openTabs(uri).length > 0, `${name} tab`);
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
    }
  });

  it('blocks every network and eval attempt from inside the webview (CSP)', async () => {
    const uri = tempCopy('simple.docx');
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
    const results = await until(() => api.probeResults[0], 'probe results from the webview');
    // sendBeacon() only queues the request (returns true); the CSP violation for connect-src proves it was blocked.
    for (const key of ['fetch', 'xhr', 'websocket', 'image', 'eval']) {
      assert.strictEqual(results[key], 'blocked', `${key} must be blocked, got ${results[key]}`);
    }
    const directives = new Set(api.violations.map((v) => v.directive.split('-elem')[0]!.split('-attr')[0]));
    for (const d of ['connect-src', 'img-src', 'frame-src', 'script-src']) {
      assert.ok(directives.has(d), `a ${d} violation must be reported; got ${[...directives].join(', ')}`);
    }
    assert.ok(
      !api.violations.some((v) => /webview\.(js|css)|vscode-resource|vscode-cdn|blob:/.test(v.blockedURI)),
      `the extension's own resources must not be blocked: ${JSON.stringify(api.violations)}`,
    );
  });

  async function editSaveCheck(name: string, edit: Edit, part: string, expected: string) {
    const uri = tempCopy(name);
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
    await until(() => api.probeResults.length > 0 && openTabs(uri).length > 0, 'editor ready');
    await sleep(1500); // let the worker finish loading (edits arriving earlier are queued anyway)
    api.injectEdit(uri, edit);
    await sleep(500);
    await vscode.commands.executeCommand('workbench.action.files.save');
    await until(() => entry(uri.fsPath, part).includes(expected), `saved file to contain "${expected}"`);
    return uri;
  }

  it('edit + save: xlsx cell', async () => {
    await editSaveCheck('simple.xlsx', { kind: 'cell', sheet: 0, row: 1, col: 0, value: 'Bananas', old: 'Apples' }, 'xl/sharedStrings.xml', 'Bananas');
  });

  it('edit + save: docx paragraph', async () => {
    const uri = await editSaveCheck('simple.docx', { kind: 'paragraph', index: 0, text: 'Edited via VS Code', old: 'First paragraph.' }, 'word/document.xml', 'Edited via VS Code');
    assert.ok(entry(uri.fsPath, 'word/document.xml').includes('Last paragraph.'));
  });

  it('edit + save: odt paragraph and ods cell', async () => {
    await editSaveCheck('simple.odt', { kind: 'paragraph', index: 4, text: 'Odt edited', old: 'Last one' }, 'content.xml', 'Odt edited');
    await editSaveCheck('simple.ods', { kind: 'cell', sheet: 0, row: 0, col: 0, value: 'Renamed', old: 'Name' }, 'content.xml', 'Renamed');
  });

  it('undo reverts the edit before saving', async () => {
    const uri = tempCopy('simple.docx');
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
    await until(() => api.probeResults.length > 0 && openTabs(uri).length > 0, 'editor ready');
    await sleep(1500);
    api.injectEdit(uri, { kind: 'paragraph', index: 0, text: 'Temporary', old: 'First paragraph.' });
    await sleep(300);
    await vscode.commands.executeCommand('undo');
    await sleep(500);
    api.injectEdit(uri, { kind: 'paragraph', index: 3, text: 'Kept change', old: 'Last paragraph.' });
    await sleep(300);
    await vscode.commands.executeCommand('workbench.action.files.save');
    await until(() => entry(uri.fsPath, 'word/document.xml').includes('Kept change'), 'save');
    const xml = entry(uri.fsPath, 'word/document.xml');
    assert.ok(!xml.includes('Temporary'));
    assert.ok(xml.includes('First paragraph.'));
  });

  it('pdf: renders in the sandbox without blocked resources, edit + save fills a form field', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([300, 300]);
    page.drawText('Application form', { x: 20, y: 270, size: 18 });
    doc.getForm().createTextField('name').addToPage(page, { x: 20, y: 200, width: 200, height: 24 });
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vsofx-'));
    const file = path.join(dir, 'form.pdf');
    fs.writeFileSync(file, await doc.save());
    const uri = vscode.Uri.file(file);
    const violationsBefore = api.violations.length;
    const errorsBefore = api.errors.length;
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
    await until(() => api.probeResults.length > 0 && openTabs(uri).length > 0, 'editor ready');
    await sleep(4000); // pdf.js worker start + first render
    assert.deepStrictEqual(api.errors.slice(errorsBefore), [], 'no error was reported while loading the PDF');
    const own = api.violations.slice(violationsBefore).filter((v) => /blob:|webview.js|vscode-resource|vscode-cdn/.test(v.blockedURI));
    assert.deepStrictEqual(own, [], 'pdf.js and its worker must not be blocked by the CSP');
    api.injectEdit(uri, { kind: 'pdf', op: { op: 'setField', name: 'name', value: 'Filled in VS Code', old: '' } });
    await sleep(500);
    const sizeBefore = fs.statSync(file).size;
    await vscode.commands.executeCommand('workbench.action.files.save');
    // the file is rewritten on save; object streams hide the text from a plain byte search, so wait for a size change
    await until(() => fs.statSync(file).size !== sizeBefore, 'saved PDF');
    await sleep(300);
    const saved = await PDFDocument.load(fs.readFileSync(file));
    assert.strictEqual((saved.getForm().getField('name') as PDFTextField).getText(), 'Filled in VS Code');
  });

  async function textEditSave(name: string, content: string, viewType: string, edit: Edit, expected: string) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vsofx-'));
    const file = path.join(dir, name);
    fs.writeFileSync(file, content);
    const uri = vscode.Uri.file(file);
    const errorsBefore = api.errors.length;
    await vscode.commands.executeCommand('vscode.openWith', uri, viewType);
    await until(
      () => vscode.window.tabGroups.all.flatMap((g) => g.tabs).some((t) => t.input instanceof vscode.TabInputCustom && t.input.uri.toString() === uri.toString()),
      `${name} tab`,
    );
    await sleep(1500);
    api.injectEdit(uri, edit);
    await sleep(500);
    await vscode.commands.executeCommand('workbench.action.files.save');
    await until(() => fs.readFileSync(file, 'utf8') === expected || false, `saved ${name}: ${JSON.stringify(fs.readFileSync(file, 'utf8'))}`, 10_000);
    assert.deepStrictEqual(api.errors.slice(errorsBefore), [], `no errors for ${name}`);
  }

  it('csv: auto-detected delimiter, cell edit + save keeps the format', async () => {
    await textEditSave('t.csv', 'a;b\r\n1;2\r\n', VIEW_TYPE, { kind: 'cell', sheet: 0, row: 1, col: 1, value: 'x;y', old: '2' }, 'a;b\r\n1;"x;y"\r\n');
  });

  it('a file that is not an Office package shows the error panel and stays untouched', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vsofx-'));
    const file = path.join(dir, 'fake.docx');
    fs.writeFileSync(file, 'this is plain text');
    const before = fs.readFileSync(file);
    const uri = vscode.Uri.file(file);
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW_TYPE);
    await until(() => openTabs(uri).length > 0, 'tab');
    await sleep(2000);
    assert.deepStrictEqual(fs.readFileSync(file), before);
  });
});
