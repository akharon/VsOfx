# VsOfx – Architecture Documentation (arc42)

Status: implemented (0.0.1), 2026-11-18. Deviations from the first draft are marked "(as built)".

## 1. Introduction and Goals

VsOfx is a small, open-source, non-commercial VS Code extension for **viewing and lightly editing Office documents** inside the editor.

| Priority | Goal |
|---|---|
| 1 | Open `.docx`, `.xlsx`, `.pptx`, `.odt`, `.ods` with a readable view. |
| 2 | Edit cell values in `.xlsx`/`.ods` and paragraph text in `.docx`/`.odt`, and save. |
| 3 | **Third-party code must never reach the network, file system or processes** (enforced, not trusted). |
| 4 | Stay small: few, well-known dependencies; own code only for simple format slices. |
| 5 | Respect every third-party license. |

Non-goals: full Office fidelity (layout, charts, macros, track changes), slide editing, `.odp`, printing, format conversion.

### Stakeholders
| Role | Expectation |
|---|---|
| Author / maintainer | Small codebase, few dependencies to review. |
| Users | Quick, safe look at Office files without leaving VS Code, simple edits. |
| Contributors | Clear structure, permissive license. |

## 2. Constraints

| Constraint | Detail |
|---|---|
| Platform | VS Code Extension API (`CustomEditorProvider` + Webview). |
| Language | TypeScript, bundled with esbuild. |
| License | Project licensed **MIT** (see 2.2). |
| Dependencies | Only permissive licenses (MIT, BSD, Apache-2.0). Runtime dependencies are allowed **only in the webview bundle**. The extension host bundle has no third-party runtime code. |
| Offline | No network access, no external binaries. |

### 2.1 Library review and selection

Reviewed candidates:

| Package | License | Verdict | Reason |
|---|---|---|---|
| `docx` 9.x | MIT | rejected | Generate only; cannot open existing files. |
| `node-xlsx` 0.25 | Apache-2.0 (npm; GitHub NOASSERTION) | rejected | Lossy save (drops styles), bundles SheetJS, unclear license metadata. |
| `nodejs-pptx` 1.2.5 | MIT | rejected | Generate only, stale, no types. |
| `officeparser` 8.x | MIT | rejected | Only slide text needed; heavy dependencies (pdfjs-dist, tesseract.js); Node ≥ 22 engine. Replaced by own pptx text extraction. |
| `odf-kit` 0.14 | Apache-2.0 | rejected | Young 0.x API, Node ≥ 22 engine; ODF XML is simple enough for own code. |

Selected (all used as npm modules, no copied code, all run **only inside the webview**):

| Package | License | Role |
|---|---|---|
| `exceljs` 4.x | MIT | Read and write `.xlsx`, keeping styles and formulas. |
| `mammoth` 1.x | BSD-2-Clause | `.docx` → HTML for viewing. |
| `fflate` 0.8.x | MIT | Zip read/write for all formats (docx patching, pptx/odt/ods parsing and saving). |
| `@xmldom/xmldom` 0.8.x | MIT | (as built) XML DOM for own format code. `DOMParser` does not exist in Web Workers; xmldom is already a dependency of mammoth, so it adds no new code to the bundle. |

Transitive: `exceljs` and `mammoth` bring `jszip`. exceljs is bundled via its prebuilt browser build (`dist/exceljs.min.js`), so its Node-only dependencies (archiver, unzipper, tmp, ...) are not in any bundle. The `uuid` advisory reported by `npm audit` concerns calls with a caller-supplied buffer, which exceljs does not make.

Note: the npm `xlsx` (SheetJS 0.18.5) package is outdated with known advisories and must not be used.

### 2.2 Licensing of VsOfx

- Project license: **MIT**, compatible with MIT and BSD-2-Clause dependencies.
- "Non-commercial" is a project intention, not a license clause. A non-commercial license (e.g. PolyForm Noncommercial) would make the project not OSI open source; MIT is recommended.
- License texts of bundled dependencies (including their transitive dependencies) ship in `THIRD_PARTY_NOTICES.md`, generated at build time.

## 3. Context and Scope

```
+--------+  open/save  +--------------------------------------------------+
| VS Code| <---------> | VsOfx                                            |
| user   |             |  Extension host (Node, trusted, own code only)   |
+--------+             |    reads/writes bytes via vscode.workspace.fs    |
                       |            ^  postMessage (bytes, edits)         |
                       |            v                                     |
                       |  Webview (browser sandbox, CSP: no network)      |
                       |    exceljs, mammoth, fflate + own format code    |
                       +--------------------------------------------------+
                                     |
                                     v
                             File system (.docx .xlsx .pptx .odt .ods)
```

Only the extension host touches the file system. Nothing touches the network.

## 4. Solution Strategy

- **Sandbox split:** the extension host is a thin, trusted byte pipe (read file, hand bytes to the webview, receive bytes, write file, manage dirty state, undo, backup). All parsing, rendering, editing and serializing happens in the **webview**, which is a browser sandbox without Node, file system or process access, and whose network access is blocked by CSP.
- One generic `CustomEditorProvider` registered for all supported extensions.
- Each format is a small **adapter** inside the webview with the same interface.
- Reuse libraries for complex formats (xlsx: exceljs, docx view: mammoth). Write own code for simple slices (pptx text, ODF text/cells, docx paragraph patching) on top of fflate.
- Save patches the original package: exceljs keeps the workbook model; docx/odt/ods edits replace only the affected XML parts in the original zip, everything else is copied unchanged.

## 5. Building Block View

```
src/
  host/                       extension host bundle (no third-party runtime deps)
    extension.ts              activate(): registers the provider
    OfficeEditorProvider.ts   CustomEditorProvider: open, save, saveAs, revert, backup, undo/redo
    OfficeDocument.ts         holds original bytes + edit list; dirty state
    webviewHtml.ts            HTML shell with strict CSP and nonce
  webview/                    webview bundle (browser)
    main.ts                   message handling, picks adapter + view
    worker.ts                 Web Worker: runs adapters off the UI thread
    protocol.ts               (as built) main thread <-> worker messages
    sanitize.ts               (as built) allowlist sanitizer; emits a new HTML string from an xmldom tree
    adapters/
      Adapter.ts              interface FormatAdapter
      zip.ts xml.ts odf.ts numFmt.ts   (as built) fflate helpers, xmldom helpers, ODF helpers, Excel number-format renderer
      xlsxAdapter.ts          exceljs
      docxAdapter.ts          view: mammoth → sanitized HTML; edit: fflate + XML patch of word/document.xml
      pptxAdapter.ts          fflate + DOMParser over ppt/slides/slideN.xml (text only, read-only)
      odtAdapter.ts           fflate + DOMParser over content.xml (paragraphs)
      odsAdapter.ts           fflate + DOMParser over content.xml (table cells)
    views/
      textView.ts             paragraphs, contenteditable per paragraph
      sheetView.ts            sheet tabs + editable grid
      slidesView.ts           slide list with text (read-only)
    style.css                 VS Code theme CSS variables
  shared/
    messages.ts               typed host ↔ webview messages
    base64.ts                 (as built) byte transport helpers
    model.ts                  DocModel = TextDocModel | SheetModel | SlidesModel; Edit types
doc/                          arch42.md, plan.md
test/                         fixtures + adapter tests
```

### FormatAdapter contract (webview)
| Method | Purpose |
|---|---|
| `load(bytes, edits)` | Parse file, replay `edits` (restoring unsaved state), return a `DocModel`. (as built) |
| `applyEdit(edit)` | Apply `{kind:'cell', sheet, row, col, value}` or `{kind:'paragraph', index, text}`. |
| `save()` | Return new file bytes: original package + edits. |
| `readonly` | `true` for pptx. |

### Host ↔ webview messages
| Direction | Message |
|---|---|
| host → webview | `init {format, bytes, edits, readonly}`, `apply {edit}` (user edit mirrored, undo, redo), `requestSave {requestId}` |
| webview → host | `ready`, `edit {edit}`, `saveResult {requestId, bytes}`, `saveError`, `error {message}`, `reopenAsText`, `securityViolation` |

(as built) `bytes` travel as **base64 strings**: VS Code serialises messages as JSON, and a Node `Buffer` arrived in the webview as `{type:"Buffer",data:[...]}` (found by the integration test). The host validates every `edit` it receives.

## 6. Runtime View

**Open:** `openCustomDocument(uri)` → host reads bytes → `resolveCustomEditor` creates webview → webview posts `ready` → host posts `init {bytes}` → worker parses with adapter → view renders. A skeleton/spinner is shown at once.

**Edit:** user edits → adapter applies locally → webview posts `edit` → host stores it, pushes undo entry, fires `onDidChangeCustomDocument` (dirty).

**Save / backup:** host posts `requestSave` → worker runs `adapter.save()` → webview posts `saved {bytes}` → host writes with `workspace.fs.writeFile`. (as built) If no editor of the document is open when saving, the host shows an error asking the user to reopen the file; reopening sends `init` with the original bytes and the edit list, the worker replays the edits, and saving works again. Any open editor of the document can serve a save (all hold the same state).

**Undo/redo:** host sends the inverse/redo edit to the webview, which applies it.

**Hidden tab:** view state is kept with `retainContextWhenHidden`; re-init replays edits if the webview was disposed.

## 7. Deployment View

- esbuild produces `dist/host.js` (Node, CommonJS, own code only) and `dist/webview.js` + `dist/worker.js` (browser).
- Packaged with `@vscode/vsce` into a `.vsix` with `LICENSE` and `THIRD_PARTY_NOTICES.md`.
- Because no Node APIs are used for parsing, the extension can also run in VS Code for the Web (vscode.dev) later.

## 8. Cross-cutting Concepts

### 8.1 Security – enforcing "no network" for untrusted libraries
1. **Isolation:** third-party code runs only in the webview/worker (no Node, no fs, no child processes).
2. **Content-Security-Policy** on the webview:
   ```
   default-src 'none'; script-src 'nonce-<n>'; style-src ${cspSource} 'unsafe-inline';
   img-src data: blob:; font-src ${cspSource}; worker-src blob: ${cspSource};
   connect-src 'none'; form-action 'none'; frame-src 'none'
   ```
   `connect-src 'none'` blocks fetch/XHR/WebSocket; restricted `img-src` prevents exfiltration via image URLs; no `'unsafe-eval'`.
3. **Webview options:** `localResourceRoots` = `dist/` only, `enableCommandUris: false`.
4. **Sanitizing:** mammoth HTML is sanitized (no scripts, event handlers or external URLs; images only as `data:`).
5. **Supply chain:** exact versions pinned, lockfile committed, `npm ci --ignore-scripts`, `npm audit` and diff review (`npm diff`) on every upgrade.
6. **Build checks (fail the build):** esbuild plugin rejects imports of `http`, `https`, `net`, `tls`, `dgram`, `dns`, `child_process`, `worker_threads`, `fs` in the webview bundle; host bundle must contain no `node_modules` code; bundle scan for `fetch(`, `XMLHttpRequest`, `WebSocket`, `sendBeacon`, `importScripts`, `eval(`, `new Function` against a reviewed allowlist.
7. **Tests:** adapters are tested with network globals stubbed to throw; optional CI run in a container with `--network none`.

7. **(as built) Worker delivery:** `dist/worker.js` is embedded as base64 in `webview.js` and started from a `blob:` URL, so no request is needed and `connect-src` stays `'none'`; the blob worker inherits the page CSP.
8. **(as built) Reporting:** the webview forwards `securitypolicyviolation` events to the host (output channel). In test builds only, `dist/probe.js` (not shipped in the `.vsix`) attempts fetch, XHR, WebSocket, image, beacon, iframe and `eval`; the integration test inside VS Code asserts all are blocked.

### 8.1a PDF in the sandbox

pdf.js runs on the webview's main thread with its own blob worker (embedded as base64, like our worker). It only ever receives bytes (`getDocument({data})`), never URLs, so its fetch/XHR code is unreachable and would be blocked by `connect-src 'none'` anyway. pdf-lib runs in our format worker. Form controls are our own HTML elements over the widget rectangles (pdf.js draws pages without form widgets), so the user never runs PDF JavaScript (`enableScripting` is not used).

### 8.2 Usability
- Parsing in a Web Worker keeps the UI responsive on large files; loading indicator.
- Lazy-load per-format code (exceljs only for spreadsheets, mammoth only for docx).
- Theming via VS Code CSS variables.
- Same file open in a second editor is read-only.

### 8.3 Error handling and data safety
- Parse error → error panel with "Reopen with text editor"; the file is never written after a failed parse.
- Save writes a package derived from the original zip; original bytes are kept for revert/backup.
- Warn before parsing files above a configurable size (default 20 MB).

## 9. Architecture Decisions

| ID | Decision | Reason |
|---|---|---|
| AD-1 | MIT license | Simple, compatible with all dependencies. |
| AD-2 | Third-party code only in the webview | Browser sandbox + CSP enforce "no network / no fs" instead of trusting libraries. |
| AD-3 | exceljs for xlsx | Complex format; keeps styles on save; MIT. |
| AD-4 | mammoth for docx viewing, own fflate-based patching for docx editing | `docx` cannot read; mammoth is read-only. |
| AD-5 | fflate as the only zip library | Small, fast, MIT; replaces jszip. |
| AD-6 | Own code for pptx text, odt and ods | Simple slices; avoids officeparser's and odf-kit's heavy/young dependencies. |
| AD-7 | Parse in a Web Worker | Responsive UI on large files. |
| AD-8 | Bytes as base64 in messages | VS Code JSON-serialises messages; binary types are not reliable. |
| AD-9 | xmldom for own XML code | No `DOMParser` in workers; already bundled via mammoth. |
| AD-10 | Own Excel number-format renderer | exceljs returns raw values; `cell.text` ignores number formats. |
| AD-11 | PDF: pdf.js (legacy build, Apache-2.0) to draw, pdf-lib fork `@cantoo/pdf-lib` (MIT) to save | pdf.js renders faithfully and lists form widgets but its own saving is tied to its viewer UI; pdf-lib fills, adds and removes AcroForm fields and edits pages. Alternatives rejected: pdfme (template generator, not a PDF editor), pdf.js viewer component (large, brings network-oriented code and CSS). The legacy build is needed because VS Code's Electron lacks newer JS APIs (`Uint8Array.toHex`). |
| AD-12 | PDF edits are absolute assignments (`src/shared/pdf.ts`) | "Field X = V", "widget W removed" ... Undo is the same op with old/new swapped, the saved file is always original bytes + state, and the host edit list, hot-exit backup and undo work unchanged. |
| AD-13 | CSV/TSV as a one-sheet table (`papaparse`, MIT, in the worker) with own delimiter detection (`,` `;` tab `\|`, most consistent column count wins) | Reuses the sheet view and cell edits. Untouched files are saved byte for byte; edited files keep delimiter, line ends, BOM and final newline, and quote only where needed. |
| AD-14 | JSON, YAML, XML, Markdown: the document is its text; one edit type `TextEdit` (splice `at/del/ins`) | Undo = swap `del`/`ins`. Tree and preview edits are turned into the smallest splice of the source, so formatting, comments and key order never change. Parsers with source offsets run in the UI thread (fast, needed on every keystroke): `jsonc-parser` (MIT), `yaml` (ISC), an own offset-keeping XML scanner (+ the browser's `DOMParser` for well-formedness), `markdown-it` (MIT, `html: false`; remote images replaced by alt text). Registered as an *option* editor (`vsofx.dataEditor`) so VS Code's text editor stays the default for these files. Text splices are posted to the host before the worker confirms them, so a following Ctrl+Z always sees them. |

## 10. Quality Requirements

| Quality | Scenario |
|---|---|
| Security | Any network attempt from third-party code fails (CSP) and is caught by tests. |
| Fidelity | Open + save of an `.xlsx`/`.docx` without edits stays openable in Excel/LibreOffice with styles intact. |
| Safety | A parse error never modifies the file on disk. |
| Performance | A 1 MB file shows content in < 2 s; UI never blocks > 100 ms during parse. |
| Size | `.vsix` under ~5 MB. |

## 11. Risks and Technical Debt

| Risk | Mitigation |
|---|---|
| exceljs browser build size. | Lazy-load; measure in Phase 2. |
| exceljs or mammoth may need Node polyfills (e.g. `Buffer`, `stream`) in the browser. | Use their browser builds; add minimal polyfills via esbuild; never add network-capable polyfills. |
| Save depends on the webview being alive. | Host keeps original bytes + edit list and can recreate the webview and replay. |
| docx text spread across several runs (`<w:r>`). | v1 replaces a paragraph's runs with one run keeping the first run's formatting. |
| Own ODF/pptx code covers text only. | Documented as non-goal; can be extended later. |
| Re-zipping recompresses entries (fflate cannot copy compressed entries). | Content of untouched parts is byte-identical after decompression; media files are stored; ODF `mimetype` is first and stored. |
| `retainContextWhenHidden` costs memory. | Acceptable for document editors; can be made a setting. |

## 12. Glossary

| Term | Meaning |
|---|---|
| OOXML | Office Open XML (`.docx`, `.xlsx`, `.pptx`). |
| ODF | OpenDocument Format (`.odt`, `.ods`). |
| Custom editor | VS Code API for a webview-based editor bound to a file type. |
| Adapter | VsOfx module (in the webview) handling one file format. |
| CSP | Content-Security-Policy, browser rules restricting what a page may load or connect to. |
