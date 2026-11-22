# VsOfx – Implementation Plan

See [arch42.md](arch42.md) for architecture, library decisions and the security model.

## Status (2026-11-22)

Version 1.3.0: Phases 0–8 are implemented (Office, ODF, PDF forms, CSV, Markdown WYSIWYG (later dropped, see below), rich text, tables, alignment, sheet selection/copy/paste, sheet add/remove/rename, column widths, wrap). The `.vsix` is 1.87 MB. **Phase 9 (below) is the current plan.** The phase descriptions after this section are kept as history; the original 0.0.1 status follows.
**Decision 2026-11-22: Markdown editor dropped.** VsOfx has no Markdown editor or viewer; VS Code's built-in Markdown editor and preview are used. Removed: the `vsofx.dataEditor` option editor, the `text` splice edit, the code view, `markdown-it`, the code highlighter and the diagram relay. All Markdown items in the phases and the backlog below are dropped.


### Status of 0.0.1 (2026-11-18)

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

---

## Phase 9 – Simple, fast, pleasant (current plan, 2026-11-22)

Aim: the everyday essentials of an office suite, done so well that people enjoy using them, in a VS Code-native, professional look, with a package **≤ 2 MB**. Not a Word/Excel clone. Architecture changes are AD-17 … AD-20 in [arch42.md](arch42.md).

### 9.1 Research summary

**What people actually use.**
- Microsoft's Word 2003 command telemetry, the only hard data found, puts Paste, Save, Copy, Undo and Bold at the top. Together they are about 32 % of all command use, and Paste alone is over 11 % (15 % in Excel). Change Font Size is #11, and after about the top 10 commands the curve is flat. ([Jensen Harris, "No Distaste for Paste"](https://learn.microsoft.com/en-us/archive/blogs/jensenh/no-distaste-for-paste-why-the-ui-part-7), [Office 2010 usage data](https://learn.microsoft.com/en-us/archive/blogs/office2010/how-does-usage-data-improve-the-office-user-experience))
- The most frequent actions are cursor movement and Backspace. Reliable caret, selection, Enter and Backspace behaviour therefore matters more than any button.
- Spreadsheets: no telemetry for individual features was found. Expert lists and Microsoft's own "Home tab" help point to editing, copy/paste, sort, filter, SUM/AutoSum, find, and the Sum/Average/Count status bar. ([Excel status bar](https://support.microsoft.com/en-US/Excel/excel-status-bar-options), [Chandoo poll](https://chandoo.org/wp/?p=5818), [Excel for the web tips](https://eus.prod.support.services.microsoft.com/en-au/office/see-tips-for-excel-for-the-web-109ca6ce-7760-4f7c-a90d-5e7e5d1f8009); weak evidence)
- An unverified figure says few users use functions beyond SUM or PivotTables. Advanced features go undiscovered because people learn spreadsheets from each other ([arXiv 1808.10642](https://arxiv.org/pdf/1808.10642)).

**What lightweight editors miss.** WordPad's typical complaints are: no table creation, no spell check or word count, and lost headers, footers and comments on save. Notepad added tables in 2025 because people asked for them. ([The Register](https://www.theregister.com/2025/11/24/notepad_tables_support/), [Windows Report](https://windowsreport.com/wordpad-windows-8-1-spell-checker/)) The worst failure is silently losing content the editor cannot show. VsOfx already avoids this (patch, don't regenerate).

**UI guidance.**
- VS Code says to theme everything with `--vscode-*` tokens, keep things accessible (contrast, ARIA, keyboard), and avoid promotions, wizards and welcome pop-ups ([webview UX](https://code.visualstudio.com/api/ux-guidelines/webviews), [UX overview](https://code.visualstudio.com/api/ux-guidelines/overview), [custom editors](https://code.visualstudio.com/api/extension-guides/custom-editors)).
- Simple editors (Google Docs, Sheets) keep one row of grouped icons and move what does not fit into a "More" menu. Known bug class to avoid: an empty or stale overflow menu after a resize ([Sheets overflow bug](https://discuss.google.dev/t/sheets-overflow-menu-ellipsis-is-broken-it-hides-menus-and-makes-them-inaccessible/280198)).
- Notion and Typora use a floating selection toolbar and a "/" insert menu ([Typora 1.14](https://support.typora.io/What's-New-1.14/)).
- Excel's status bar shows Sum, Average and Count for a selection.

**Competitors (VS Code Marketplace).**

| Extension | Scope | Notes |
|---|---|---|
| [Office Viewer](https://marketplace.visualstudio.com/items?itemName=cweijan.vscode-office) | Office, PDF, Markdown, archives, HTTP, … | 1.7 M installs; broad, unfocused |
| [Excel Viewer](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer) | xlsx/csv grid | 6.9 M installs, no longer maintained; keeps per-file view state |
| [Edit CSV](https://marketplace.visualstudio.com/items?itemName=janisdd.vscode-edit-csv) | CSV editing | 2.6 M installs; known save/rename issues, slow on big files |
| [Rainbow CSV](https://marketplace.visualstudio.com/items?itemName=mechatroner.rainbow-csv) | CSV colouring in the text editor | 25 M installs; light |

What we can offer that they don't: one focused, small, sandboxed editor with VS Code-native save, undo and hot exit, the same toolbar in every format, and editing (not only viewing) for Word, Writer, Excel and Calc.

### 9.2 Principles for every change
1. **Top-10 first.** Per family, the ten most-used actions must work with one click and with the standard shortcut. Nothing from the long tail gets added until they do.
2. **Same action, same place.** Same icon, tooltip, shortcut and toolbar group in every format (arch42 §8.2).
3. **No surprise, no loss.** Every edit is one undo step; untouched content is saved unchanged; if something can't be done, the tooltip says why.
4. **Pay for size.** Each feature states its bundle cost; the package stays ≤ 2 MB (gate 1.9 MB).
5. **Calm UI.** No welcome pages, no pop-ups, no settings beyond the two existing ones unless one is needed.

### 9.3 Ideas and feasibility

Value: H/M/L. Effort: S ≈ ≤ 1 day, M ≈ 2–3 days, L ≈ a week. Size: change of the packaged `.vsix`.

| ID | Idea | Value | Effort | Risk | Size | Verdict |
|---|---|---|---|---|---|---|
| **P1** | Embed worker sources as string literals, not base64 (AD-17) | H (room for everything else, faster start) | S | low | **−480 KB** | Wave 1 |
| **P2** | Load pdf.js + PDF view + pdf.js worker as a chunk only for PDFs (AD-17) | H (non-PDF files open faster) | M | medium (start-up path) | ±0 | Wave 1 |
| **P3** | Package size gate (fail > 1.9 MB) and a size line in the build output | M | S | none | 0 | Wave 1 |
| **D1** | Word/Writer: one editing surface. Select across paragraphs, Enter splits, Backspace/Delete merges, delete or type over a multi-paragraph selection, paste several lines. New `delpara` edit (AD-18) | **H** (top user request; caret/Backspace are the most frequent actions) | L | medium-high | +5 KB | Wave 1 |
| **S1** | Sheets: status line with Sum, Average, Count (numeric cells) and Count (all) for the selection | H | S | low | +2 KB | Wave 1 |
| **T1** | Toolbar overflow: buttons that don't fit go into a "More" (…) menu, order stays stable, never an empty menu (AD-20) | H (split editors, laptops) | M | low | +3 KB | Wave 1 |
| D2 | Paragraph style drop-down (Normal, Heading 1–3, Quote) instead of the cycling heading button; Word, Writer | H | S | low | +1 KB | Wave 2 |
| D3 | Shortcuts everywhere: Ctrl+B/I/U, Ctrl+K link, Ctrl+Shift+7/8 lists, Ctrl+Alt+0..3 styles, Ctrl+Shift+V paste as plain text; shown in tooltips; Ctrl+/ shows the list | H | S | low (VS Code chords: test) | +2 KB | Wave 2 |
| D4 | Underline and strikethrough buttons (Word/Writer already support the runs) | M | S | low | +0.5 KB | Wave 2 |
| D5 | Word count, sheet/cell position in VS Code's own status bar (host `StatusBarItem`) | M | S | low | +1 KB | Wave 2 |
| D6 | Remember zoom, active sheet, scroll per file (`vscode.setState`) | M | S | low | +1 KB | Wave 2 |
| S2 | Sort the selected range or the column A→Z / Z→A (value cells; refuse with a message when it holds formulas or merges) as one `batch` | H | M | medium | +3 KB | Wave 2 |
| S3 | Insert / delete row and column: CSV and ODS fully; xlsx only when the sheet has no formulas, merges or defined names that would need shifting (else explain) | H | M-L | medium-high | +4 KB | Wave 2 |
| S4 | AutoSum (Alt+=): writes `=SUM(range)` below/right of the selection; the own formula engine shows the result | M | S | low | +1 KB | Wave 2 |
| N1 | "VsOfx: New Document" command: blank .docx/.xlsx/.odt/.ods/.csv from the command palette and explorer | M (makes VsOfx a tool, not only a viewer) | M | low | +6 KB | Wave 2 |
| P4 | Bytes as transferred `Uint8Array` instead of base64 (AD-19) | M (big files) | M | medium (needs integration test) | 0 | Wave 3 |
| P5 | Per-family worker chunks (sheets: exceljs/papaparse; documents: mammoth; PDF: pdf-lib) | M (start-up) | M | medium | ±0 | Wave 3 |
| U1 | Loading skeleton instead of "Loading…", clearer read-only banner for pptx/odp/odg | M | S | low | +1 KB | Wave 3 |
| U2 | Floating selection toolbar (B I U link) in Word/Writer | M | M | medium (focus handling) | +3 KB | Wave 3, only if D3 is not enough |
| X1 | Slide text editing (pptx/odp) | M | L | high | +10 KB | Later; needs a design note |
| — | Cut/Copy/Paste **buttons** | Paste is #1, but webviews cannot read the clipboard from a button without a permission prompt | — | — | — | Not doing; Ctrl+C/V/X and the context menu work |
| — | Spell check | Webviews only have the browser's checker, which is off on purpose (noise in code-like text) | — | — | — | Not doing; maybe a setting later |
| — | Ribbon, track changes, comments editing, charts, pivots, mail merge, macros, page layout, headers/footers editing, collaboration, AI | Long tail / outside the aim / size | — | — | — | Not doing |

**Size check of the plan:** 1.87 MB − 480 KB (P1) + about 35 KB (all of Wave 1 and Wave 2) ≈ **1.43 MB**, which leaves about 470 KB of headroom under the 1.9 MB gate.

### 9.4 Wave 1 (started 2026-11-22: two Sonnet agents and one Haiku agent)

**Result (2026-11-22):** P1–P3, D1, S1 and T1 are implemented. `npm run check` is green (255 tests in 20 files). The `.vsix` is 1.31 MB (was 1.87 MB): `webview.js` is 2.1 MB instead of 5.0 MB, and pdf.js is in `dist/pdf.js`, loaded only for PDFs. All of this is verified by automated tests only, not yet by hand in VS Code (checklist in the 1.4.0 changelog entry). Known limits of D1: a selection that crosses a table or an image is ignored. Backspace at the start of a list item merges it into the previous block instead of first removing the bullet. A removed paragraph that holds a section break is emptied instead of removed. Toggle buttons in the "More" menu keep `aria-pressed` (should become `menuitemcheckbox`).

Every agent works only in its own files, runs only its own tests, and does not commit. The integrator (main session) then runs `npm run check`, `npm run package`, reviews the diff and reports.

| Agent | Tasks | Files it owns |
|---|---|---|
| Sonnet A | P1, P2, P3 | `esbuild.mjs`, `src/webview/main.ts` (worker start-up, PDF case), `src/webview/virtual.d.ts`, `src/host/webviewHtml.ts`, `src/host/OfficeEditorProvider.ts` (only the HTML/URI part), `scripts/package.mjs`, `scripts/check-bundles.mjs`, `.vscodeignore`, `test/worker-bundle.test.ts`, `test/host.test.ts` |
| Sonnet B | D1 | `src/webview/views/textView.ts`, `src/shared/model.ts` (`delpara` + `isEdit`/`inverseEdit`), `src/webview/adapters/docxAdapter.ts`, `docxWrite.ts`, `odtAdapter.ts`, `src/host/OfficeEditorProvider.ts` (only `editLabel`), `src/webview/code.css` (`.page` rules), new `test/singleSurface.test.ts` |
| Haiku C | S1, T1 | `src/webview/views/sheetView.ts` (status line only), `src/webview/ui/toolbar.ts`, `src/webview/style.css`, `test/toolbar.test.ts`, new `test/sheetStatus.test.ts` |

**Done when (Wave 1):**
- P1–P3: `npm run package` reports the size and the `.vsix` is ≤ 1.5 MB. A PDF still opens, renders and saves. A non-PDF document does not load the PDF chunk. The CSP is unchanged.
- D1: in jsdom tests, for Word and Writer: select from paragraph 2 into paragraph 4 and type "x" → one undo step that leaves paragraph 2 = prefix + "x" + suffix of paragraph 4, with 3 and 4 removed; undo restores all three. Backspace at the start of a paragraph joins it to the one before; Delete at the end joins the next one. Enter in the middle splits. Pasting 3 lines makes 3 paragraphs. Save + reload shows the result in docx and odt. Existing textView tests stay green.
- S1: selecting 2+ cells shows "Sum … Average … Count …" (numbers only for Sum/Average; formatted like the cells' locale); one cell shows nothing; it updates on edit.
- T1: at a width where buttons don't fit, a "More" button appears and holds exactly the hidden buttons (same click handlers, titles, pressed state); widening restores them; never an empty menu; keyboard reachable (Enter/Escape/arrows).

### 9.5 Wave 2 and 3
Wave 2: D2, D3, D4, D5, D6, S2, S3, S4, N1, in this order (cheap, high-value first). Wave 3: P4, P5, U1, then decide on U2. Each item gets its "Done when" line before work starts.

Done (Wave 3): P4, P5, U1; U2 skipped (D3 shortcuts are enough for now).
- P4: init and save messages carry `Uint8Array` (`shared/bytes.ts` `asBytes`/`ownedBytes`); main thread -> worker transfers; worker -> main thread and main thread -> host transfer the saved bytes. Backups unchanged (host-side `OfficeDocument.serialize`). Test: `test/transfer.test.ts`. Not verified inside VS Code by hand: the integration suite has not been run for this change.
- P5 (measured with esbuild metafiles, minified): single worker 1942 KiB (webview.js 2020 KiB with it embedded) -> webview.js 71 KiB; worker per family: sheet 1078 KiB (exceljs 923), doc 378 KiB, pdf 593 KiB. gzip of all shipped chunks about 1237 KB -> 1287 KB (+50 KB, +4 %). Worth it: Word/Writer/slides parse 81 % less, spreadsheets 45 % less. Test: `test/worker-bundle.test.ts` (each bundle holds only its libraries). Limit: the sheet worker also holds exceljs for CSV (CSV does not need it).
- U1: skeleton in the page HTML (`shared/skeleton.ts`, styles in `style.css`), grid for spreadsheets, page otherwise; the error state is unchanged. The pptx/odp/odg read-only banner is unchanged.

Done: S2, S3, S4 (sort as a batch of cell edits; `lines` edit for insert/delete with formula rewrites as cell edits in the same batch, merges refused; AutoSum for xlsx/ods). Limits: see CHANGELOG 1.4.0.

Done: D2, D3, D4, D5 (word count only; sheet stats stay in the sheet view), D6, N1. D2 offers Normal, Heading 1-3 and Quote (no Title). D3 chords: Ctrl+B/I/U/K, Ctrl+Shift+X, Ctrl+Shift+7/8, Ctrl+Alt+0..3 (by key position); not done: Ctrl+Shift+V and the Ctrl+/ list. D4 also fixed mammoth dropping underline on open (`styleMap: u => u`). D6 saves zoom, scroll of scrollable areas (by CSS class) and the active sheet tab through the webview state; a PDF's scroll may not be restored because its pages render late. N1: command palette only, no CSV, no Explorer menu; the host writes a stored (uncompressed) zip itself (`src/host/newDocument.ts`, about 3 KB per file).

## Milestones

| Version | Content |
|---|---|
| 0.0.1 | Phases 0–1: sandboxed skeleton, safety rails, CI |
| 0.1.0 | + Phase 2: xlsx view/edit |
| 0.2.0 | + Phase 3: docx view/edit |
| 0.3.0 | + Phases 4–5: pptx view, odt/ods view/edit |
| 1.0.0 | + Phase 6: hardened, published |
| 1.0 – 1.3 | Phases 7–8: toolbar, rich text, PDF forms, CSV, Markdown (dropped again in 1.4.0), tables, alignment, sheet essentials |
| 1.4.0 | Phase 9 Wave 1: smaller and faster package, one editing surface for Word/Writer, sheet status line, toolbar overflow |
| 1.5.0 | Phase 9 Wave 2: styles drop-down, shortcuts, sort, insert/delete rows/columns, AutoSum, New Document |

## Open questions
- `.xls` (binary Excel 97–2003): not supported. Options: a read-only view through an own BIFF reader (large) or "Open with text editor / convert in Excel" (current). Decide before Wave 3.
- Whether the floating selection toolbar (U2) is wanted after D3 shortcuts are in.
- Closed: license MIT; `engines.vscode ^1.90`; mammoth kept (docx-preview not needed); `@js-preview/*` not used.
