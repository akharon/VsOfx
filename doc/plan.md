# VsOfx – Implementation Plan

See [arch42.md](arch42.md) for architecture, library decisions and the security model.

## Status (2026-11-18)

Phases 0–6 are implemented as version 0.0.1. `npm run check` is green (lint, typecheck, build, bundle scan, 99 unit tests, license check), and `npm run test:integration` passes 7 tests inside a real VS Code (all formats open; fetch, XHR, WebSocket, image, iframe and eval are blocked by the CSP; edit, save and undo for xlsx, docx, odt, ods). `npm run package` produces `vsofx-0.0.1.vsix` (about 0.6 MB).

Deviations from this plan (details in arch42.md, "as built"):
- **xmldom instead of DOMParser** in the worker (`DOMParser` is not available there). The sanitizer walks an xmldom tree and emits a new HTML string.
- **Bytes as base64** in host/webview messages (found by the integration test).
- **Number formats:** own renderer (`numFmt.ts`) because exceljs does not format cell text.
- **Fixtures are generated** (`test/helpers/fixtures.ts`, `npm run fixtures`), not created in Excel/LibreOffice. The 50k-row sheet is generated inside the test.
- **Zip re-writing** recompresses entries (fflate has no raw copy); the content of untouched parts is identical.
- **No `@vscode/test-cli` download in CI by default:** CI uses `xvfb-run`; locally set `VSCODE_EXECUTABLE_PATH`.
- **Not done / manual:** opening saved files in real Excel, Word and LibreOffice (checklist in CONTRIBUTING.md); marketplace publishing (6.4); an extension icon; a repository URL in package.json.
- Open questions below: `docx-preview` and `@js-preview/*` were not evaluated; mammoth and own views were used.

- Runtime dependencies (webview only, exact versions pinned): `exceljs`, `mammoth`, `fflate`.
- Extension host: own code only, no third-party runtime dependencies.
- Each task has a **Done when** line. A phase is finished when all its tasks are done and `npm run check` passes.

---

## Phase 0 – Project setup and safety rails

### 0.1 Repository
- `git init`; add `.gitignore` (`node_modules/`, `dist/`, `*.vsix`, `.vscode-test/`).
- `LICENSE` (MIT, year 2026, author name).
- `README.md`: purpose, supported formats, security model summary (link to arch42 §8.1), build instructions.
- **Done when:** first commit contains these files.

### 0.2 Scaffold
- `package.json`:
  - `name: vsofx`, `displayName: VsOfx`, `license: MIT`, `engines.vscode: ^1.90.0` (to confirm, see open questions).
  - `main: ./dist/host.js`, `browser: ./dist/host.js` (prepares vscode.dev support).
  - `activationEvents`: none needed (custom editors activate automatically).
  - `contributes.customEditors`: one entry `vsofx.officeEditor`, `selector` for `*.docx`, `*.xlsx`, `*.pptx`, `*.odt`, `*.ods`, `priority: default`.
- Dev dependencies: `typescript`, `esbuild`, `@types/vscode`, `@vscode/vsce`, `@vscode/test-electron`, `@vscode/test-cli`, `vitest`, `eslint` + `typescript-eslint`, `license-checker-rseidelsohn`.
- `tsconfig.json` (strict, `ES2022`), with `src/host` using `lib: ["ES2022"]` and `src/webview` using `lib: ["ES2022","DOM","WebWorker"]` (two tsconfig files extending a base).
- Folder layout as in arch42 §5 (`src/host`, `src/webview`, `src/shared`, `test/`).
- **Done when:** `npm install` and `npm run build` succeed on an empty project.

### 0.3 Build (`esbuild.mjs`)
- Three bundles:
  | Entry | Output | Platform | Format |
  |---|---|---|---|
  | `src/host/extension.ts` | `dist/host.js` | node | cjs, `external: ['vscode']` |
  | `src/webview/main.ts` | `dist/webview.js` | browser | iife |
  | `src/webview/worker.ts` | `dist/worker.js` | browser | iife |
- `--watch` mode for development, sourcemaps in dev, minify in production.
- npm scripts: `build`, `watch`, `package` (vsce), `test`, `lint`, `check` (= lint + typecheck + build checks + tests + license check).
- **Done when:** `npm run build` produces the three files.

### 0.4 Supply-chain rules
- `.npmrc`: `ignore-scripts=true`, `save-exact=true`.
- Commit `package-lock.json`; CI uses `npm ci`.
- `CONTRIBUTING.md` section "Upgrading a dependency": run `npm diff --diff=<pkg>@old --diff=<pkg>@new`, `npm audit`, review, then update.
- **Done when:** a fresh `npm ci` runs no install scripts.

### 0.5 Build checks (`scripts/check-bundles.mjs`)
- esbuild plugin in the webview/worker builds: fail on import of `http`, `https`, `net`, `tls`, `dgram`, `dns`, `child_process`, `worker_threads`, `fs`, `os`.
- Host bundle check: use esbuild `metafile`; fail if any input path contains `node_modules`.
- Bundle scan: search `dist/*.js` for `fetch(`, `XMLHttpRequest`, `WebSocket`, `EventSource`, `sendBeacon`, `importScripts`, `eval(`, `new Function`. Every hit must be listed in `scripts/allowed-hits.json` (file, pattern, count, reason). New hits fail the build.
- **Done when:** adding `fetch('x')` to a webview file makes `npm run check` fail.

### 0.6 License notices
- `scripts/notices.mjs`: run `license-checker-rseidelsohn --production --json`, fail on any license not in the allowlist (`MIT`, `ISC`, `BSD-2-Clause`, `BSD-3-Clause`, `Apache-2.0`, `0BSD`), write `THIRD_PARTY_NOTICES.md` with full license texts.
- **Done when:** the file is generated and included in `.vsix` (checked in Phase 6).

### 0.7 CI
- `.github/workflows/ci.yml`: Node LTS, `npm ci`, `npm run check`, VS Code integration tests with `xvfb-run` on Linux.
- **Done when:** CI is green on the first push.

---

## Phase 1 – Sandboxed editor skeleton

Starting point: Microsoft `vscode-extension-samples/custom-editor-sample` (MIT), the binary "paw draw" editor, adapted. Keep its copyright notice in copied files and add it to `THIRD_PARTY_NOTICES.md`.

### 1.1 Shared types (`src/shared/`)
- `model.ts`:
  ```ts
  type Format = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'ods';
  type Edit =
    | { kind: 'cell'; sheet: number; row: number; col: number; value: string; old: string }
    | { kind: 'paragraph'; index: number; text: string; old: string };
  ```
  Each edit carries `old` so the host can build the inverse for undo without parsing.
- `messages.ts`: discriminated unions
  - host → webview: `init {format, bytes: Uint8Array, edits: Edit[], readonly}`, `apply {edit}`, `requestSave {requestId}`
  - webview → host: `ready`, `edit {edit}`, `saveResult {requestId, bytes} | saveError {requestId, message}`, `error {message}`
- **Done when:** both bundles compile against these types.

### 1.2 `OfficeDocument` (`src/host/OfficeDocument.ts`)
- Implements `vscode.CustomDocument`.
- State: `uri`, `originalBytes`, `edits: Edit[]` (applied list), `savedEditCount`.
- `static async create(uri, backupId?)`: read bytes with `vscode.workspace.fs.readFile` (from backup if `backupId` is set; the backup stores original bytes + JSON edit list).
- Emits `onDidChange` with `{ label, undo, redo }` for each edit.
- **Done when:** unit test creates a document, adds edits, undo/redo changes `edits` correctly.

### 1.3 `OfficeEditorProvider` (`src/host/OfficeEditorProvider.ts`)
- Implements `vscode.CustomEditorProvider<OfficeDocument>`, registered with `supportsMultipleEditorsPerDocument: true` and `webviewOptions.retainContextWhenHidden: true`.
- `resolveCustomEditor`: set webview options (`enableScripts: true`, `enableCommandUris: false`, `localResourceRoots: [dist]`), set HTML, wait for `ready`, send `init`. The first webview of a document is the **primary** (editable); others get `readonly: true`.
- `saveCustomDocument` / `saveCustomDocumentAs`: send `requestSave` to the primary webview, await `saveResult` (timeout 30 s → error), write bytes, set `savedEditCount`. If no webview is alive (edge case): show an error asking the user to reopen the editor; reopening sends `init` with `originalBytes` + `edits`, the worker replays the edits, then save works.
- `revertCustomDocument`: reload `originalBytes` from disk, clear edits, re-send `init`.
- `backupCustomDocument`: write `{edits}` JSON + original bytes to `context.destination`.
- Undo/redo: send `apply` with the inverse/redo edit to all webviews of the document.
- **Done when:** manual test: open, edit (fake adapter), Ctrl+Z, Ctrl+S, revert all behave correctly.

### 1.4 Webview HTML (`src/host/webviewHtml.ts`)
- Generates nonce, builds CSP exactly as in arch42 §8.1, loads `dist/webview.js` and `dist/style.css` via `asWebviewUri`.
- Passes the worker URL in a `data-worker` attribute. The worker is loaded as a `blob:` URL created from fetched script text (fetch of the local `vscode-resource` is allowed only if needed; preferred: inline the worker bundle as text into `webview.js` at build time, so no fetch is required and `connect-src` stays `'none'`).
- **Done when:** DevTools of the webview shows no CSP violations for our own code.

### 1.5 Webview runtime (`src/webview/main.ts`, `worker.ts`)
- `main.ts`: `acquireVsCodeApi()`, post `ready`, on `init` → show spinner → forward bytes (transferable) to the worker → render returned model with the matching view.
- `worker.ts`: holds the active adapter, handles `load`, `applyEdit`, `save` requests.
- Views send user edits to `main.ts` → worker `applyEdit` → host `edit`.
- Error panel with message and a button that posts `error`; host offers "Reopen with Text Editor" (`vscode.openWith` + `default`).
- **Done when:** a `fakeAdapter` (bytes shown as hex, one editable field) round-trips through open/edit/save.

### 1.6 Security tests
- Integration test: inject a script call `fetch('https://example.com')` in a test build of the webview; assert a CSP violation event is reported and no request succeeds.
- **Done when:** test runs in CI.

### 1.7 Format detection
- `src/webview/adapters/index.ts`: map file extension → lazy `import()` of the adapter (esbuild code splitting is not available for iife; instead keep one worker bundle but initialise libraries lazily; measure size in 2.4).
- Verify zip signature (`PK\x03\x04`) before parsing; otherwise show error.
- **Done when:** a renamed `.txt` → `.docx` shows the error panel and the file is untouched.

---

## Phase 2 – Spreadsheets (`.xlsx`) with exceljs

### 2.1 `xlsxAdapter`
- `load`: `new ExcelJS.Workbook().xlsx.load(bytes)`; build `SheetModel { sheets: [{ name, rows: CellView[][], merges, colWidths }] }` with `CellView { text, style? }`. Use `cell.text` for display (formatted value); formulas show the cached result, the formula is available in a tooltip.
- `applyEdit(cell)`: number-looking input → number, `=`-prefixed → formula (no evaluation, result cleared), otherwise string.
- `save`: `workbook.xlsx.writeBuffer()`.
- Limit: show at most 10,000 rows × 200 columns per sheet; show a notice beyond that.
- **Done when:** vitest round-trip: load fixture, edit B2, save, reload → B2 changed, styles of other cells unchanged.

### 2.2 `sheetView`
- Sheet tabs at the bottom, HTML `<table>` with sticky header row (A, B, …) and row numbers.
- Double-click or Enter to edit a cell (`<input>` overlay); Enter/Tab commits, Esc cancels.
- Basic styles: bold, italic, fill colour, font colour, alignment, column widths, merged cells.
- Virtualised rendering (only visible rows) when a sheet has more than 1,000 rows.
- **Done when:** manual test with `test/fixtures/styles.xlsx` looks reasonable in light and dark theme.

### 2.3 Fixtures and tests
- `test/fixtures/`: `simple.xlsx`, `styles.xlsx`, `formulas.xlsx`, `multi-sheet.xlsx`, `large.xlsx` (50k rows), all created with LibreOffice / Excel and committed.
- Tests: load each, round-trip without edits opens in LibreOffice (manual checklist), edit round-trip (automated).

### 2.4 Bundle size and browser compatibility
- Measure `dist/worker.js`; if exceljs needs Node polyfills (`buffer`, `stream`, `events`), add them via esbuild `alias` using browser polyfill packages (licenses must pass 0.6; none may be network-capable).
- **Done when:** `.xlsx` works in the webview with no Node globals, bundle size recorded in README.

---

## Phase 3 – Word (`.docx`)

### 3.1 Viewing with mammoth
- `docxAdapter.load`: `mammoth.convertToHtml({ arrayBuffer }, { convertImage: images as data: URIs })`.
- Sanitize the HTML with own allowlist sanitizer (`src/webview/sanitize.ts`, uses `DOMParser`): allowed tags `p h1–h6 strong em u s sub sup ul ol li table thead tbody tr th td a img br span`; allowed attributes `href` (only `#…`), `src` (only `data:image/`), `colspan`, `rowspan`; drop everything else.
- **Done when:** fixtures render; sanitizer unit tests reject `<script>`, `onerror=`, `javascript:` and external `src`.

### 3.2 Paragraph mapping
- In parallel to mammoth, parse `word/document.xml` (fflate `unzipSync` + `DOMParser`) and list body paragraphs `w:body/w:p` (including those inside tables) in document order.
- Map mammoth output blocks to paragraph indices by order and normalised text equality; paragraphs that cannot be mapped are shown read-only.
- **Done when:** for all fixtures at least 95 % of text paragraphs are mapped (test reports the ratio).

### 3.3 Editing and saving
- `textView`: each mapped block is `contenteditable` (plain text only, paste as text). Commit on blur or Enter → `paragraph` edit.
- `docxAdapter.save`: for each edited paragraph, keep `w:pPr`, keep the first `w:r`'s `w:rPr`, remove all other runs, write one `w:r` with `w:t xml:space="preserve"`. Serialise with `XMLSerializer`, re-zip with fflate `zipSync` copying all other entries byte-for-byte (keep entry order and compression).
- **Done when:** round-trip test: edit a paragraph, save, reload → new text present; file opens in Word/LibreOffice (manual checklist).

### 3.4 Fixtures
- `simple.docx`, `styles.docx` (headings, lists), `tables.docx`, `images.docx`, `split-runs.docx` (one paragraph with mixed formatting).

---

## Phase 4 – Presentations (`.pptx`, read-only)

### 4.1 `pptxAdapter`
- Unzip; read `ppt/presentation.xml` → `p:sldIdLst` order; resolve slide paths via `ppt/_rels/presentation.xml.rels`.
- For each slide: collect text from `a:p` (join `a:t` per paragraph) per shape `p:sp`, keep title shape (`p:ph type="title"|"ctrTitle"`) separately; notes from `ppt/notesSlides/` if present.
- Optionally first image per slide (`a:blip r:embed` → media as `data:` URI).
- **Done when:** unit test on `simple.pptx` returns slides in the right order with correct titles.

### 4.2 `slidesView`
- Vertical list of slide cards: number, title, bullet text, optional image, collapsible notes. Read-only banner.
- **Done when:** manual test on fixtures.

---

## Phase 5 – OpenDocument (`.odt`, `.ods`)

### 5.1 Common
- Unzip; check `mimetype` entry (`application/vnd.oasis.opendocument.text` / `.spreadsheet`).
- On save: `mimetype` must be the **first entry, stored uncompressed** (ODF requirement); copy all other entries unchanged; update only `content.xml`.

### 5.2 `odtAdapter`
- Parse `content.xml`, list `text:p` and `text:h` in `office:text` (including table cells) in order; text from child nodes, handling `text:s` (spaces), `text:tab`, `text:line-break`, `text:span`.
- Edit: replace the paragraph's children with one text node, keeping the element and its `text:style-name`.
- Reuse `textView` (headings rendered by `text:h` outline level).
- **Done when:** round-trip test passes; file opens in LibreOffice.

### 5.3 `odsAdapter`
- Parse `table:table` → rows/cells, expanding `table:number-rows-repeated` / `table:number-columns-repeated` (cap trailing empty repeats).
- Display value: `text:p` content; type from `office:value-type`.
- Edit: write `office:value-type` + `office:value` for numbers, `text:p` for display; split repeated cells when editing one inside a repeated range.
- Reuse `sheetView` (no styles in v1).
- **Done when:** round-trip test passes, including an edit inside a repeated range.

---

## Phase 6 – Hardening and release

### 6.1 Tests
- Unit tests (vitest) for all adapters and the sanitizer, with `fetch`, `XMLHttpRequest`, `WebSocket` replaced by functions that throw (vitest `setupFiles`).
- Integration tests (`@vscode/test-cli`): open each fixture, edit, save, undo, revert.
- Optional CI job: run tests in Docker with `--network none`.

### 6.2 UX polish
- Settings: `vsofx.maxFileSizeMB` (default 20), `vsofx.retainContextWhenHidden` (default true).
- Status bar item: format and "read-only" when applicable.
- Command `VsOfx: Reopen with Text Editor`.
- Light/dark/high-contrast theme check.

### 6.3 Packaging
- `.vscodeignore`: exclude `src/`, `test/`, `node_modules/`, maps; include `dist/`, `LICENSE`, `THIRD_PARTY_NOTICES.md`, `README.md`, icon.
- `vsce package`; verify `.vsix` size < 5 MB and contents (`vsce ls`).
- `CHANGELOG.md` for 0.1.0.

### 6.4 Publishing (optional)
- Marketplace publisher account and Open VSX namespace; `vsce publish`, `ovsx publish`.

---

## Phase 7 – Toolbar and Office power features

Source: the "Word + Excel power user cheat sheet" the project owner supplied. Every feature is sorted by what it costs in this architecture. Rules that do not change: no network, all third-party code runs in the worker (arch42 §8), the host stays a byte pipe, and the file is only changed through `Edit` objects that carry their `old` value (so undo/redo and the "restore the original" rule keep working).

### 7.1 Done (toolbar v1: view-only, no file format change)
A button bar (`src/webview/ui/toolbar.ts`) sits above every document; text labels only, no icon font.

| Where | Button | Cheat-sheet item |
|---|---|---|
| all editable | Undo, Redo, Save (ask the host to run the standard VS Code command; whitelist `EDITOR_COMMANDS`) | AutoSave/Save, Ctrl+Z |
| all | Find (Ctrl+F; Enter / Shift+Enter cycle; uses the CSS Custom Highlight API, no DOM changes) | Find |
| Word, slides | Zoom A− / 100 % / A+ | – |
| Word | Outline (navigation pane built from the headings) | Navigation Pane |
| Word | ¶ (formatting marks at paragraph ends) | Show Formatting Marks |
| Word | word count | – |
| Excel | name box + formula bar (read-only) | Formula auditing (partly) |
| Excel | Freeze row (first row sticky) | Freeze Panes (view only, not saved) |
| Excel | Filter (text filter per column; the first row is the header; view only, sheets up to 1000 rows) | Filters, Ctrl+Shift+L |

Tests: `test/toolbar.test.ts` (jsdom), integration suite unchanged and green.

### 7.2 Next – small, view-only (each is about a day)
- **Sheet zoom.** Needs the virtual scroller to scale `ROW_HEIGHT`; do it with a CSS variable plus a recalculation, not CSS `zoom` (it breaks scroll maths).
- **Range selection** (Shift+arrows, drag) with a status line showing Sum / Count / Average. Prerequisite for copy and fill.
- **Copy / paste as TSV** (Ctrl+C on a range; paste a block as a batch of cell edits, see 7.3).
- **Keyboard**: Ctrl+Arrow (jump to data edge), Ctrl+Home / Ctrl+End, Page Up/Down.
- **Sort (view only)** and filters for virtualised sheets: needs a row permutation layer between the model and the DOM; edits must still use the model's row index.
- **Named ranges** and **data validation** shown (exceljs exposes `definedNames` and `dataValidations`): name box drop-down, validation list as a drop-down in the cell editor, still unenforced.
- **Conditional formatting** rendering for the simple rule types (cell value, colour scale); everything else is ignored and says so.
- **Document outline for sheets** (sheet list with row counts) and a sheet-tab context menu (rename is a file change, see 7.3).

### 7.3 Then – file-changing edits (new `Edit` kinds, each with tests that the rest of the package stays byte-identical)
All of these need a `batch` edit kind (`{ kind: 'batch', edits: Edit[] }`) so that one user action is one undo step. Add it to `model.ts`, `isEdit`, `inverseEdit` (reverse order) and the host's `OfficeDocument` first.
- **Fill down / right** (Ctrl+D, Ctrl+R), **Remove duplicates** (clears cells), **Text to columns**: all expressible as batches of cell edits.
- **Autosum** (Alt+=): writes a formula cell; exceljs does not calculate, so the cached value is shown as empty until the file is opened in Excel. Decide whether to compute simple SUM/AVERAGE/COUNT ourselves (small own evaluator; a general formula engine is out of scope, HyperFormula is GPL/commercial and therefore not allowed).
- **Insert / delete row or column** (xlsx): must shift merges, formulas, defined names, conditional formats; high risk, do it only for sheets without these, or refuse with a clear message.
- **Freeze panes saved** into the file (`sheetViews/pane`).
- **Word: apply a paragraph style** (Heading 1..3, Normal): new `style` edit that sets `w:pStyle`; the style list comes from `styles.xml`.
- **Word: insert / delete paragraph.** The current paragraph mapping is a sequence alignment between mammoth's output and `document.xml`; inserting shifts indices, so first give every paragraph a stable id from its position in `document.xml`.
- **Word: bold / italic / underline of a whole paragraph** (run properties), then of a selection (splits runs).
- **ODF equivalents** of the above, after the xlsx/docx versions are stable.

### 7.4 Later – larger or uncertain (each needs a design note before work starts)
- **Word: Track Changes** display (`w:ins` / `w:del`, `comments.xml`) as a read-only layer first; writing tracked changes later.
- **Word: Compare** two files or a file against git HEAD (paragraph diff in the webview; VS Code's own diff cannot show binary files).
- **Word: TOC / cross references / section breaks** shown as stored fields (we never recalculate them; Word updates them on open). Say so in the UI.
- **Excel Tables (Ctrl+T)**, **PivotTables**, **charts**, **sparklines**: display of the stored result only (charts: render the cached series with a small own SVG drawer); creating them is not planned.
- **Power Query, XLOOKUP and other formula calculation, Flash Fill, Quick Analysis, Watch Window**: need a formula engine or are Excel-only features; not planned, listed so the question is closed.
- **Restrict editing / workbook protection**: display the "protected" flag and make the view read-only for protected sheets; never try to break protection.
- **AutoSave**: VS Code's own `files.autoSave` already drives the custom editor's save. Verify with an integration test (set `files.autoSave: afterDelay`, inject an edit, wait for the file to change), then mention it in the README.

---

## Phase 8 – PDF viewing and form editing

Done: open `.pdf` (pdf.js render + selectable text + Find), fill text/checkbox/radio/dropdown/list fields, add text field/checkbox/dropdown, move/resize/delete any form field, add fixed text, rotate and delete pages, full undo/redo, save through pdf-lib (untouched files are saved byte for byte).

Next:
- Add radio groups, list boxes, date and signature-image fields; field properties (required, read-only, max length, default value, tab order).
- Real annotations: highlight, ink, stamp, comments (pdf.js annotation editors write them).
- Reorder, insert and merge pages; extract pages.
- Unicode form values and free text (embed a Unicode font; today only WinAnsi/Latin text can be saved).
- Passwords and encryption, flatten form, digital signatures.
- Page thumbnails, outline/bookmarks, CJK and JPEG2000 support (needs pdf.js CMaps/wasm shipped inside the VSIX).
- Edit existing page text (hard in PDF; likely never).

## Milestones

| Version | Content |
|---|---|
| 0.0.1 | Phases 0–1: sandboxed skeleton, safety rails, CI |
| 0.1.0 | + Phase 2: xlsx view/edit |
| 0.2.0 | + Phase 3: docx view/edit |
| 0.3.0 | + Phases 4–5: pptx view, odt/ods view/edit |
| 1.0.0 | + Phase 6: hardened, published |

## Open questions
- Confirm the project license: MIT (recommended) or a non-commercial license (not OSI open source).
- Minimum supported VS Code version (`engines.vscode`).
- Evaluate `docx-preview` (Apache-2.0) as a higher-fidelity alternative to mammoth for viewing (would add `jszip`); decide before Phase 3.
- Evaluate `@js-preview/*` viewers (license and source availability unverified) before Phases 2–4.
