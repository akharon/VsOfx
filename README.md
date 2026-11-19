# VsOfx – Office files in VS Code

View and lightly edit Office documents and PDF forms without leaving the editor. Open source (MIT), non-commercial in spirit.

| Format | View | Edit |
|---|---|---|
| `.xlsx` | sheets, styles, merged cells, formulas (cached results) | cell values, formulas as text input |
| `.docx` | headings, lists, tables, images | paragraph text |
| `.odt` | headings, lists, tables, bold/italic/underline | paragraph text |
| `.ods` | cell values (formulas show cached results) | cell values |
| `.pptx` | slide text, notes, first picture per slide | read-only |
| `.pdf` | pages, selectable text, form fields | fill and add form fields, move/resize/delete fields, add text, rotate/delete pages |

Editing replaces a paragraph's text and keeps the paragraph style and the formatting of its first run. Other runs of that paragraph lose their individual formatting. Everything else in the file is saved unchanged.

## Toolbar

A button bar sits above each document:

- **All:** Undo, Redo, Save, Find (Ctrl+F; Enter and Shift+Enter move between matches), zoom (documents and slides).
- **Word and ODT:** Outline (headings), ¶ (formatting marks), word count.
- **Excel and ODS:** name box and formula bar, Freeze row, Filter (view only, sheets up to 1,000 rows; the first row is the header). Freeze and filter are not saved into the file.

More features are planned; see `doc/plan.md`, Phase 7.

## PDF forms

Open a `.pdf` and fill the fields directly on the page (text, checkboxes, radio buttons, dropdowns). Press **Design** to edit the form itself: add a text field, checkbox or dropdown (click the button, then the page), drag to move, use the corner handle to resize, and press Delete or the Delete button to remove any field. **＋ Text** prints fixed text on the page; **Rotate page** and **Delete page** change the pages. Everything supports Undo/Redo, and the file is only written when you save. Pages are drawn with pdf.js, saving is done with pdf-lib; the PDF's JavaScript never runs.

Limits: only Latin (WinAnsi) characters can be saved into fields and printed text; PDFs with non-embedded CJK fonts or JPEG 2000 images may render incompletely; password-protected PDFs are not supported yet. See `doc/plan.md`, Phase 8.

## Security: no network, enforced

Third-party libraries (`exceljs`, `mammoth`, `fflate`, `@xmldom/xmldom`, `pdfjs-dist`, `@cantoo/pdf-lib`) run **only inside a VS Code webview worker**, never in the extension host:

- The webview has a strict Content-Security-Policy: `default-src 'none'`, `connect-src 'none'`, no `eval`, scripts only with a nonce.
- The extension host contains only our own code (read/write bytes, undo, backup).
- `npm run check` fails if a bundle imports Node I/O modules, if new `fetch`/`XMLHttpRequest`/`eval`-like code appears, or if a license is not allowed. Tests run with all network functions replaced by throwing stubs, and an integration test inside VS Code verifies that fetch, XHR, WebSocket, images, iframes and `eval` are blocked.

Details: `doc/arch42.md` §8.1.

## Build

Requires Node.js >= 22.18. Run each command on its own line: Windows PowerShell 5.1 does not support `&&` (use `;` or separate lines).

```
npm ci --ignore-scripts
npm run build        # dist/
npm run check        # lint, typecheck, build, bundle scan, tests, license check
npm run test:integration   # runs inside VS Code (set VSCODE_EXECUTABLE_PATH to use an installed one)
npm run package      # creates vsofx-0.<commits>.0-<commit>.vsix (version and name come from git)
```

In VS Code: *Run and Debug* → "Run Extension", or *Terminal → Run Build Task* → "Package VSIX". Install the result with `code --install-extension vsofx-<version>-<commit>.vsix`.

Sizes (production): `worker.js` ≈ 1.2 MB (mostly exceljs), `webview.js` ≈ 1.7 MB (includes the worker as base64).

## Settings

- `vsofx.maxFileSizeMB` (default 20): ask before opening larger files.
- `vsofx.retainContextWhenHidden` (default true): keep the editor state in hidden tabs (needs a window reload).

Command: *VsOfx: Reopen with Text Editor*.

## Limits

No charts, macros, track changes or page layout. Spreadsheets show at most 10,000 rows × 200 columns. Typed input in `.ods` that starts with `=` is stored as text. Dates in `.xlsx` are shown with common number formats only.

## License

MIT, see `LICENSE`. Third-party licenses: `THIRD_PARTY_NOTICES.md`.
