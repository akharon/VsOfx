# VsOfx – Architecture Documentation (arc42)

Status: as built in 1.4.0 (2026-11-22). Decisions marked **(planned)** belong to Phase 9 in [plan.md](plan.md) and are not implemented yet.

## 1. Introduction and Goals

VsOfx is a small, open-source VS Code extension for **viewing and editing everyday documents inside the editor**: Word and Writer documents, Excel, Calc and CSV sheets, slides (read-only) and PDF forms. It covers the essentials of an office suite, not all of it.

| Priority | Goal |
|---|---|
| 1 | **Safe:** third-party code can never reach the network, the file system or processes (enforced by the sandbox, not trusted). |
| 2 | **Simple and pleasant:** the everyday 20 % of Word/Excel features, reachable in one click or one familiar shortcut, in a VS Code-native look. |
| 3 | **Faithful:** saving changes only what the user edited; everything else in the file stays as it was. |
| 4 | **Small and fast:** packaged `.vsix` **under 2 MB**; a document shows content quickly and the UI never freezes. |
| 5 | Respect every third-party license (permissive only). |

Non-goals: full Office fidelity (page layout, headers/footers, charts, pivots, macros, track changes), slide editing, formula recalculation beyond the own small engine, real-time collaboration, cloud services, printing, format conversion.

### Stakeholders
| Role | Expectation |
|---|---|
| Users | Open an Office/PDF/CSV file in VS Code and do the common edits without leaving the editor; it looks and behaves like VS Code. |
| Author / maintainer | Small codebase, few dependencies to review, tests for every edit kind. |
| Contributors | Clear structure, one place per format, permissive license. |

## 2. Constraints

| Constraint | Detail |
|---|---|
| Platform | VS Code Extension API (`CustomEditorProvider` + Webview), `engines.vscode ^1.90`. Desktop and (prepared) VS Code for the Web. |
| Language / build | TypeScript (strict), bundled with esbuild. Node ≥ 22.18 for the build only; `scripts/npm.ps1` provides a verified portable Node when none is installed. |
| Size | Marketplace package **≤ 2 MB** (build gate at 1.9 MB, see §10). |
| License | Project **MIT**. Dependencies: MIT, ISC, BSD, Apache-2.0, 0BSD only (checked by `scripts/notices.mjs`). |
| Dependencies | Runtime dependencies are allowed **only in the webview / worker bundles**. The host bundle contains own code only. |
| Offline | No network access, no external binaries, no downloads at runtime. |

### 2.1 Libraries (all run only in the webview or its workers)

| Package | License | Role | Share of bundle (min.) |
|---|---|---|---|
| `exceljs` 4.4 (browser build) | MIT | Read/write `.xlsx` with styles and formulas | ~945 KB (worker) |
| `pdfjs-dist` 6.5 (legacy build) | Apache-2.0 | Draw PDF pages, text layer, form widget list | ~500 KB + 1.3 MB own worker |
| `@cantoo/pdf-lib` 2.11 | MIT | Save PDF form values, fields, pages, images | ~460 KB (worker) |
| `mammoth` 1.13 (+ jszip, underscore, lop …) | BSD-2-Clause | `.docx` → HTML for viewing | ~300 KB incl. dependencies |
| `@xmldom/xmldom` 0.8 | MIT | XML DOM in the worker (no `DOMParser` there) | ~70 KB |
| `papaparse` 5.5 | MIT | CSV/TSV parsing | ~20 KB |
| `fflate` 0.8 | MIT | Zip read/write for all own format code | ~14 KB |

Rejected (unchanged reasons): `docx` (generate only), `node-xlsx`/SheetJS npm 0.18 (outdated, advisories), `officeparser` and `odf-kit` (heavy or young), HyperFormula (GPL/commercial), `@vscode/webview-ui-toolkit` (deprecated since 2025-01).

## 3. Context and Scope

```
+--------+  open/save  +-----------------------------------------------------------+
| VS Code| <---------> | VsOfx                                                     |
| user   |             |  Extension host (Node, trusted, own code only)            |
+--------+             |    vscode.workspace.fs, edit list, undo/redo, backups     |
                       |            ^  postMessage (bytes, edits, commands)        |
                       |            v                                              |
                       |  Webview UI thread (CSP: no network, no eval)             |
                       |    views, toolbar, pdf.js renderer           |
                       |            ^  postMessage                                 |
                       |            v                                              |
                       |  Format worker (blob:)        pdf.js worker (blob:)       |
                       |    adapters, exceljs, mammoth, pdf-lib, papaparse         |
                       +-----------------------------------------------------------+
                                     |
                                     v
     File system (.docx .xlsx .pptx .odt .ods .odp .odg .pdf .csv .tsv)
```

Only the extension host touches the file system. Nothing touches the network. The host also opens web/mail links in the user's browser on request (`openLink`, allowlisted schemes).

## 4. Solution Strategy

- **Sandbox split.** The host is a thin, trusted byte pipe: read the file, hand the bytes to the webview, keep the list of edits, run undo/redo, write the bytes it gets back. Parsing, rendering, editing and saving all happen in the webview and its workers.
- **One editor, many adapters.** One `CustomEditorProvider` (`vsofx.officeEditor`, default for Office/PDF/CSV). Each format is an adapter with the same interface; each document family has one view.
- **Edits are data.** Every change is an `Edit` object that carries its old value. Undo is the inverse edit, the saved file is "original bytes + edits", and hot-exit backups store the original bytes plus the edit list. One user action is one edit (a `batch` when it touches several things), so it is one undo step.
- **Patch, don't regenerate.** Saving rewrites only the XML elements that edits touched; all other zip entries are copied unchanged. Untouched files are saved byte for byte.
- **Libraries for the hard formats, own code for simple slices.** exceljs (xlsx), mammoth (docx view), pdf.js + pdf-lib (PDF); own code for docx writing, ODF, pptx/odp text, CSV dialect, formulas and number formats.
- **(planned) Load only what a document needs.** Code and data that only one format family uses (pdf.js, its worker) are separate files loaded on demand, and no bundle is stored base64-encoded (§8.5).

## 5. Building Block View

```
src/
  host/                         extension host bundle (own code only)
    extension.ts                activate(): providers, commands
    OfficeEditorProvider.ts     open, save, saveAs, revert, backup, undo/redo, edit labels, link relay
    OfficeDocument.ts           original bytes + edit list, saved position, dirty state
    webviewHtml.ts              HTML shell, CSP with nonce
  shared/                       used by host, webview and worker
    model.ts                    Format, DocModel, Edit union, isEdit (validation), inverseEdit, applyTextEdit
    richtext.ts                 ParaFmt (runs, block kind, alignment), virtual paragraph indices
    pdf.ts                      PDF edit ops (absolute assignments)
    formula.ts                  small spreadsheet formula engine (SUM, AVERAGE, IF, …)
    messages.ts                 host <-> webview messages
    base64.ts                   byte transport helpers
  webview/
    main.ts                     UI entry: starts workers, routes messages, builds toolbar + view
    workers/                    format workers: core.ts (protocol loop) + sheet.ts, doc.ts, pdf.ts (one per family)
    protocol.ts, family.ts      worker request/response protocol; format -> worker family
    sanitize.ts                 allowlist HTML sanitizer (worker side)
    adapters/                   one per format (worker side)
      Adapter.ts index.ts       FormatAdapter contract, format → adapter
      xlsxAdapter.ts            exceljs
      csvAdapter.ts             papaparse + own dialect detection and writer
      docxAdapter.ts docxWrite.ts  mammoth view + own paragraph/table/run writer
      odtAdapter.ts odsAdapter.ts odf.ts   own ODF text and table code
      pptxAdapter.ts odpAdapter.ts         slide text (read-only)
      pdfAdapter.ts             pdf-lib
      numFmt.ts textCodec.ts source.ts zip.ts xml.ts   helpers
    views/                      UI thread
      textView.ts               Word / Writer page
      sheetView.ts              grid, selection, sheet tabs, formula bar
      slidesView.ts             slide cards (read-only)
      pdfView.ts                pages, form overlay, PDF tools
    ui/                         toolbar.ts (buttons, icons, zoom), findBar.ts
    style.css code.css          theme-variable styling
doc/                            arch42.md, plan.md
scripts/                        npm.ps1, package.mjs, notices.mjs, check-bundles.mjs, fixtures
test/                           vitest unit tests (jsdom), integration tests in VS Code
```

### 5.1 FormatAdapter contract (worker)
| Member | Purpose |
|---|---|
| `load(bytes, edits)` | Parse, replay `edits` (unsaved state after reopen or hot exit), return a `DocModel`. |
| `applyEdit(edit)` | Apply one edit; returns an `ApplyResult` (e.g. the recalculated cell) for the view. |
| `save()` | New file bytes: original package + edits. |
| `readonly` | `true` for pptx, odp, odg. |

### 5.2 Models and edits
`DocModel = TextModel | SheetModel | SlidesModel | PdfModel | CodeModel`.

| Edit kind | Formats | Meaning |
|---|---|---|
| `cell` | xlsx, ods, csv | value of one cell (formula when it starts with `=`) |
| `colwidth`, `wrap`, `align` | xlsx, ods | column width, wrap text, horizontal alignment of a cell |
| `sheet` | xlsx, ods | add, remove, rename a sheet |
| `paragraph` | docx, odt | text and (rich) `fmt` runs, block kind, alignment of one paragraph |
| `newpara` | docx, odt | a paragraph the user added (Enter), or its removal |
| `table` | docx, odt | a table inserted after a paragraph, or removed |
| `delpara` **(planned)** | docx, odt | an existing paragraph removed (merge, multi-paragraph delete), or restored |
| `pdf` | pdf | absolute assignment ops (field value, widget moved, page rotated …) |
| `batch` | all | several edits as one undo step |

The host validates every received edit with `isEdit` before storing it.

### 5.3 Host ↔ webview messages
| Direction | Messages |
|---|---|
| host → webview | `init {format, bytes, edits, readonly}`, `apply {edit}` (undo, redo, mirror to other editors), `requestSave {requestId}` |
| webview → host | `ready`, `edit {edit}`, `saveResult {requestId, bytes}`, `saveError`, `error`, `reopenAsText`, `openLink {href}`, `command {undo\|redo\|save}`, `securityViolation` |

`bytes` are `Uint8Array`s (AD-19); receivers read them with `asBytes()`, which also accepts the old base64 string.

## 6. Runtime View

**Open:** `openCustomDocument` reads the bytes (or a backup) → `resolveCustomEditor` sets the HTML → webview posts `ready` → host posts `init` → UI thread starts the format worker and sends the bytes (transferred) → adapter parses and replays edits → view renders. PDF: the UI thread also starts the pdf.js worker and keeps its own copy of the bytes.

**Edit:** the view turns a user action into one `Edit` → worker applies it → the view shows the result → the webview posts `edit` → host validates, stores it, fires `onDidChangeCustomDocument` (dirty, undo label).

**Save / backup:** host posts `requestSave` → view flushes pending typing → worker `save()` → `saveResult` → host writes with `workspace.fs.writeFile`. Any open editor of the document can serve the save. With no editor open, the host asks the user to reopen; reopening replays the edits.

**Undo / redo:** host sends the inverse or the redo edit as `apply` to every webview of the document.

**Revert / change on disk:** host re-sends `init`; the view is closed without sending its pending input.

## 7. Deployment View

| File | Content |
|---|---|
| `dist/host.js` | host (node, cjs, own code) |
| `dist/webview.js` + `dist/webview.css` | UI thread (about 71 KB, no libraries). |
| `dist/worker-sheet.js`, `dist/worker-doc.js` | classic scripts, loaded on demand with the page nonce; each sets `globalThis.vsofxWorkers[family]` to the source (string literal) of that family's format worker (sheet: xlsx, ods, csv; doc: docx, odt, pptx, odp). `dist/worker-<family>.raw.js` is the intermediate worker and is not shipped. |
| **(planned)** `dist/pdf.js` | pdf.js, the PDF view and the pdf.js worker source, loaded only for PDF files |

The `.vsix` (built by `npm run package`, named from git) contains `dist/` bundles, `package.json`, `README.md`, `CHANGELOG.md`, `LICENSE`, `THIRD_PARTY_NOTICES.md` and the icon. The workers start from `blob:` URLs, because VS Code serves resources from another origin than the webview, and `new Worker(url)` needs the same origin.

## 8. Cross-cutting Concepts

### 8.1 Security – enforcing "no network" for untrusted libraries
1. **Isolation:** third-party code runs only in the webview and its workers (no Node, fs or processes).
2. **Content-Security-Policy** on the webview:
   ```
   default-src 'none'; script-src 'nonce-<n>'; style-src ${cspSource} 'unsafe-inline';
   img-src data: blob:; font-src ${cspSource}; worker-src blob: ${cspSource};
   connect-src 'none'; form-action 'none'; frame-src 'none'
   ```
   `connect-src 'none'` blocks fetch/XHR/WebSocket/EventSource; `img-src` without remote origins prevents exfiltration through images; no `'unsafe-eval'`. `blob:` workers inherit this policy. **(planned)** A lazily loaded chunk is a `<script>` from `dist/` that carries the same nonce; the policy does not change.
3. **Webview options:** `localResourceRoots` = `dist/` only, `enableCommandUris: false`.
4. **Sanitising:** mammoth and own HTML go through the allowlist sanitizer (no scripts, event handlers or external URLs; images only as `data:`). Links are kept inert (`data-href`) and opened only through the host for `http(s)`/`mailto`.
5. **PDF:** pdf.js only ever gets bytes, never URLs; PDF JavaScript is never run (`enableScripting` off); form controls are own HTML elements.
6. **Supply chain:** exact versions, committed lockfile, `ignore-scripts`, `npm audit`, `npm diff` review on upgrades, Dependabot, CodeQL.
7. **Build checks (fail the build):** forbidden Node-module imports in browser bundles; host bundle without `node_modules` code; bundle scan for `fetch(`, `XMLHttpRequest`, `WebSocket`, `sendBeacon`, `importScripts`, `eval(`, `new Function` against a reviewed allowlist.
8. **Tests:** unit tests run with network globals stubbed to throw. In test builds, `dist/probe.js` (not shipped) tries fetch, XHR, WebSocket, image, beacon, iframe and `eval`; the integration test asserts that all are blocked. `securitypolicyviolation` events are logged by the host.

### 8.2 UI and usability concepts
- **VS Code-native look:** only theme CSS variables (`--vscode-*`), 16 px stroke icons in the codicon style drawn inline as SVG (no icon font, no remote resource), focus rings from `--vscode-focusBorder`, works in light, dark and high-contrast themes.
- **One toolbar contract for all formats:** `[Undo Redo Save] | format group | insert group | spacer | Find, Zoom, view toggles, Source`. The same action has the same icon, tooltip text and shortcut in every format. **(planned)** Buttons that do not fit move into a "More" menu instead of wrapping (AD-20).
- **Keyboard first:** the standard shortcuts work and are shown in tooltips (Ctrl+B/I/U, Ctrl+K, Ctrl+F, Ctrl+Z/Y/S, Ctrl+± zoom).
- **Responsiveness:** parsing and saving run in the worker; large sheets are virtualised; a loading state is shown at once.
- **Honest limits:** a read-only format says so; a button that cannot work for the current element explains why in its tooltip instead of failing silently.

### 8.3 Text documents: editing surface
Today every paragraph is its own `contenteditable` (`data-p` = paragraph index). Selection therefore ends at a paragraph border, and Backspace at the start of a paragraph does not join it to the previous one.

**(planned, AD-18)** The page becomes one editing surface. Typing inside a paragraph stays a `paragraph` edit. Structural input is intercepted in `beforeinput` and turned into one `batch`: Enter splits (`paragraph` + `newpara`), Backspace/Delete at an edge merges (`paragraph` + `delpara`/`newpara remove`), deleting or typing over a selection across paragraphs keeps the first, edits it, and removes the others, and pasting several lines inserts paragraphs. Paragraph indices stay stable (they refer to the original document order), so undo and replay keep working.

### 8.4 Error handling and data safety
- Parse error → error panel with "Open with Text Editor"; a file that failed to parse is never written.
- Saving builds from the original bytes; untouched files are saved byte for byte; the ODF `mimetype` entry stays first and stored.
- Ask before parsing files above `vsofx.maxFileSizeMB` (default 20).
- An edit that the worker cannot apply shows an error and is not sent to the host.

### 8.5 Size and load-time budget
- The packaged `.vsix` must stay **≤ 2 MB**; `npm run package` fails above 1.9 MB.
- **(planned, AD-17)** Worker sources are embedded as JavaScript string literals, not base64. Measured on 1.3.0: the format worker compresses to 651 KB instead of 944 KB, the pdf.js worker to 390 KB instead of 577 KB. That saves about 480 KB in the package, and there is no `atob` copy at start-up.
- **(planned)** pdf.js and the PDF view are loaded only for PDF files, so the other formats parse about 3 MB less JavaScript when they open.
- New dependencies need a size note in this file (bundle KB before/after).

## 9. Architecture Decisions

| ID | Decision | Reason |
|---|---|---|
| AD-1 | MIT license | Simple, compatible with all dependencies. |
| AD-2 | Third-party code only in the webview / workers | Browser sandbox + CSP enforce "no network / no fs" instead of trusting libraries. |
| AD-3 | exceljs for xlsx | Complex format; keeps styles on save. |
| AD-4 | mammoth for docx viewing, own writer for docx editing | `docx` cannot read; mammoth is read-only. |
| AD-5 | fflate for all own zip code | Small, fast. (exceljs and mammoth still bring jszip internally.) |
| AD-6 | Own code for pptx/odp text, odt and ods | Simple slices; avoids heavy or young dependencies. |
| AD-7 | Parse in a Web Worker | Responsive UI on large files. |
| AD-8 | Bytes as base64 in host messages | A Node `Buffer` arrived as `{type:'Buffer',data:[…]}`. Superseded by AD-19. |
| AD-9 | xmldom for own XML code | No `DOMParser` in workers; already a mammoth dependency. |
| AD-10 | Own Excel number-format renderer | exceljs returns raw values. |
| AD-11 | PDF: pdf.js (legacy build) to draw, `@cantoo/pdf-lib` to save | Faithful rendering; pdf-lib edits AcroForms and pages. VS Code's Electron needs the legacy build. |
| AD-12 | PDF edits are absolute assignments | Undo = same op with old/new swapped; saved file = original + state. |
| AD-13 | CSV/TSV as a one-sheet table with own delimiter detection | Reuses the sheet view; untouched files saved byte for byte; dialect kept. |
| AD-14 | (dropped 2026-11-22) Markdown editor | VS Code's built-in Markdown editor and preview are enough; removing it also drops markdown-it and the code highlighter. |
| AD-15 | Word/Writer rich edits carry `fmt` runs + block kind + alignment; tables are `table` edits with virtual cell indices | Only edited paragraphs are rewritten; existing styles and numbering stay. |
| AD-16 | Alignment: horizontal only (`w:jc`, derived ODF styles) | Vertical alignment was removed: no use case in the target audience. |
| AD-17 **(planned)** | Embed worker sources as string literals; load pdf.js as an on-demand chunk | −480 KB package, ~3 MB less start-up JS for non-PDF files; CSP unchanged. |
| AD-18 **(planned)** | One `contenteditable` page for Word/Writer, structural input via `beforeinput`, new `delpara` edit | Selection across paragraphs, Enter/Backspace/Delete/paste behave like a word processor; edits stay small and stable. |
| AD-19 | Transfer bytes as `Uint8Array` instead of base64 | VS Code ≥ 1.57 transfers binary data in webview messages; saves 33 % and a copy on big files. The worker also receives and returns them as transferables. `asBytes()` stays tolerant (base64 string, serialised Buffer, number array). |
| AD-21 | One format worker per family | Measured: the single worker was 1942 KiB (exceljs 923, pdf-lib 448); now sheet 1078, doc 378, pdf 593. +50 KB gzip in the package for about 80 % less code parsed for Word/Writer/slides. A blob: worker cannot `importScripts` under the CSP, so each family is a separate worker source. |
| AD-20 **(planned)** | Toolbar overflow menu | Narrow editor groups (split view) keep a one-row toolbar. |

## 10. Quality Requirements

| Quality | Scenario |
|---|---|
| Security | Any network attempt from third-party code fails (CSP) and is caught by tests. |
| Fidelity | Open + save without edits keeps the file byte for byte; with edits it opens in Word / Excel / LibreOffice with untouched content intact. |
| Safety | A parse error never modifies the file on disk; every edit kind has an undo test. |
| Usability | The ten most-used actions per family (§8.2, plan Phase 9) work with one click or their standard shortcut, in every format of that family. |
| Performance | A 1 MB file shows content in < 2 s; UI never blocks > 100 ms during parse; typing in a 200-page document has no visible delay. |
| Size | `.vsix` ≤ 2 MB (build gate 1.9 MB). |

## 11. Risks and Technical Debt

| Risk | Mitigation |
|---|---|
| Package size close to the limit (1.87 MB in 1.3.0). | AD-17; size gate; size note for new dependencies. |
| exceljs is the largest dependency (~945 KB). | Keep; an own xlsx writer is possible later (ods shows the scale) but risky for fidelity. |
| Webview start parses all bundles for every format. | AD-17 lazy pdf chunk; later per-family worker chunks. |
| Single editing surface (AD-18) can break edit mapping. | `beforeinput` intercepts all structural changes; a fallback re-sync compares blocks by `data-p`; dedicated jsdom tests per input type. |
| Save depends on a live webview. | Host keeps original bytes + edits and replays on reopen. |
| Re-zipping recompresses entries. | Untouched part content is identical; media stored; ODF `mimetype` first and stored. |
| Shortcuts can collide with VS Code keybindings (e.g. Ctrl+K chords). | Handle keys in the webview and `preventDefault`; test in the integration suite. |
| `retainContextWhenHidden` costs memory. | Setting `vsofx.retainContextWhenHidden`. |

## 12. Glossary

| Term | Meaning |
|---|---|
| OOXML | Office Open XML (`.docx`, `.xlsx`, `.pptx`). |
| ODF | OpenDocument Format (`.odt`, `.ods`, `.odp`, `.odg`). |
| Custom editor | VS Code API for a webview-based editor bound to a file type. |
| Adapter | Worker-side module handling one file format. |
| Edit | Serializable description of one change, with its old value; the unit of undo. |
| Chunk | A separately loaded script file in `dist/`, started only when a format needs it. |
| CSP | Content-Security-Policy: browser rules on what a page may load or connect to. |
