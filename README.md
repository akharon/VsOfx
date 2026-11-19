# VsOfx – Office files in VS Code

View and lightly edit Office documents and PDF forms without leaving the editor. Open source (MIT), non-commercial in spirit.

| Format | View | Edit |
|---|---|---|
| `.xlsx` | sheets, styles, merged cells, formulas | cell values and formulas (`=SUM(A1:A3)`); dependent formulas are recalculated |
| `.docx` | headings, lists, tables, images, links | paragraph text and formatting: bold, italic, code, links, headings, bullet/numbered lists, quotes, task boxes; insert tables |
| `.odt` | headings, lists, tables, bold/italic/underline | paragraph text |
| `.ods` | cell values, formulas (shown in Excel syntax) | cell values and formulas, recalculated like `.xlsx` |
| `.pptx` `.odp` | slide text, notes, first picture per slide | read-only |
| `.odg` | page text, first picture per page | read-only |
| `.pdf` | pages, selectable text, form fields | fill and add form fields, move/resize/delete fields, add text, rotate/delete pages |
| `.csv` `.tsv` | table (delimiter `,` `;` tab `\|` detected automatically) | cell values; delimiter, quoting, line ends and encoding are kept |
| `.md` | edited like a Word document (WYSIWYG); **Source** shows the Markdown text | formatted text, headings, lists, task boxes, quotes, code, links, tables; or the Markdown source |

CSV opens in VsOfx by default. Markdown keeps VS Code's text editor as default: use **VsOfx: Open in Visual Editor** (editor title button or Explorer context menu) or *Reopen Editor With… → VsOfx Markdown Editor*. To make it the default, add `"workbench.editorAssociations": { "*.md": "vsofx.dataEditor" }`.

Editing replaces a paragraph's text and keeps the paragraph style and the formatting of its first run. Other runs of that paragraph lose their individual formatting. Everything else in the file is saved unchanged.

## Toolbar

A button bar sits above each document:

- **All:** **Source** (right edge) shows the file as text: Markdown source (editable), CSV text, the XML parts of Office/OpenDocument packages and the raw PDF (read-only, including unsaved edits). Undo, Redo, Save, Find (Ctrl+F; Enter and Shift+Enter move between matches), zoom (documents and slides).
- **Word and ODT:** Outline (headings), ¶ (formatting marks), word count.
- **Word only:** Bold, Italic (Ctrl+B / Ctrl+I), Code, Heading (cycles H1 → H2 → H3 → normal), bulleted / numbered list, task box (☐ → ☑ → none), Quote, Link (Ctrl+K), insert a 3×3 table below. Shift+Enter is a line break inside a paragraph; pasting inserts plain text.
- **Excel and ODS:** name box and formula bar, Freeze row, Filter (view only, sheets up to 1,000 rows; the first row is the header). Freeze and filter are not saved into the file.

More features are planned; see `doc/plan.md`, Phase 7.

## PDF forms

Open a `.pdf` and fill the fields directly on the page (text, checkboxes, radio buttons, dropdowns). Press **Design** to edit the form itself: add a text field, checkbox or dropdown (click the button, then the page), drag to move, use the corner handle to resize, and press Delete or the Delete button to remove any field. **＋ Text** prints fixed text on the page; **Rotate page** and **Delete page** change the pages. Everything supports Undo/Redo, and the file is only written when you save. Pages are drawn with pdf.js, saving is done with pdf-lib; the PDF's JavaScript never runs.

Limits: only Latin (WinAnsi) characters can be saved into fields and printed text; PDFs with non-embedded CJK fonts or JPEG 2000 images may render incompletely; password-protected PDFs are not supported yet. See `doc/plan.md`, Phase 8.

## Security: no network, enforced

Third-party libraries (`exceljs`, `mammoth`, `fflate`, `@xmldom/xmldom`, `pdfjs-dist`, `@cantoo/pdf-lib`, `papaparse`, `markdown-it`) run **only inside the VS Code webview** (worker or webview page), never in the extension host:

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

No charts, macros, track changes or page layout. Spreadsheets show at most 10,000 rows × 200 columns. Formulas are recalculated by a built-in engine (arithmetic, comparisons, `&`, ranges, other-sheet references and common functions such as SUM, AVERAGE, MIN, MAX, COUNT(A/IF), SUMIF, IF, IFERROR, AND/OR, ROUND, VLOOKUP, INDEX/MATCH, text functions); formulas using other functions keep the result saved in the file, and Excel recalculates everything when the file is opened. Dates in `.xlsx` are shown with common number formats only.

## License

MIT, see `LICENSE`. Third-party licenses: `THIRD_PARTY_NOTICES.md`.
