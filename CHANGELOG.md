# Changelog

## 1.4.0 (unreleased)
- Word and Writer: the page is one editing area. Select across paragraphs and type or delete; Backspace/Delete join paragraphs; Enter splits; pasting several lines makes several paragraphs. Each action is one undo step.
- Spreadsheets: Sum, Average and Count of the selected cells next to the sheet tabs.
- Toolbar: buttons that do not fit move into a "More" menu instead of wrapping.
- Smaller and faster: the package is 1.24 MB instead of 1.87 MB, and the PDF engine loads only for PDF files.
- Markdown editor removed. Use VS Code's built-in Markdown editor and preview (the `.md` files open in the normal text editor again).
- Spreadsheets: sort A→Z / Z→A by the active column (numbers as numbers; one undo step; not over formulas or merged cells); AutoSum (Σ or Alt+=); insert and delete rows and columns from the toolbar or by right-clicking a row or column header (xlsx, ods, csv). Formulas that point at moved cells are rewritten (also from other sheets); deleted references become #REF!. Not available for sheets with merged cells. Defined names, conditional formatting, data validation and charts are not moved (xlsx); sorting moves values, not cell formatting; undoing a deleted ODS column restores values but not their formatting.
- Source button removed from the toolbar. Use "Reopen Editor With… → Text Editor" to see the raw file.
- Toolbar: toggles inside the "More" menu are announced as checkable menu items (menuitemcheckbox, aria-checked).
- Word and Writer: paragraph style drop-down (Normal, Heading 1-3, Quote; over several selected paragraphs it is one undo step) replaces the cycling heading button. Underline and strikethrough buttons. Shortcuts shown in the tooltips: Ctrl+B / I / U, Ctrl+Shift+X (strikethrough), Ctrl+K (link), Ctrl+Shift+7 / 8 (numbered / bulleted list), Ctrl+Alt+0..3 (Normal, Heading 1..3).
- Word documents now show underlined text (it was dropped when opening).
- Word count of Word and Writer documents in VS Code's status bar.
- VsOfx remembers zoom, scroll position and the active sheet of each open file (also after a window reload).
- New command "VsOfx: New Document": empty Word (.docx), Spreadsheet (.xlsx), Writer (.odt) or Calc (.ods) file, saved where you choose and opened in VsOfx. Not offered: CSV and the Explorer context menu.

- Faster start: the page script is 71 KB instead of 2 MB. Each kind of file loads only its own engine (Word, Writer and slides: 378 KB; spreadsheets and CSV: 1.1 MB; PDF: its own chunk), so a Word document no longer parses the spreadsheet and PDF libraries. The package grows by about 50 KB.
- File bytes now travel between VS Code, the editor page and the worker as binary data (no base64) and are handed over instead of copied, which speeds up opening and saving big files.
- Loading shows a calm placeholder (toolbar and page or grid, with theme colours; no animation when the system asks for reduced motion) instead of the word "Loading…", from the first moment.

## 1.3.0 (unreleased)
- Marketplace and GitHub: icon, repository links, categories; release workflow (GitHub Release, VS Code Marketplace, Open VSX), CodeQL, Dependabot, issue and PR templates, security policy.
- PDF: large inserted images are reduced (at most 2400 px on a side, about 3 MB) so undo and backups stay small.
- Error messages from the editor are shortened in notifications and the log.

## 1.2.0
- PDF: insert an image from a file or by pasting; move and resize it (Shift resizes freely).
- PDF: font size, bold, italic and colour for new text fields and fixed text; multi-line text fields; fixed text wraps.
- Markdown: highlighted code blocks (Python, C, C++, Java, Rust, JavaScript, TypeScript, Dart, Ruby, GDScript) and config blocks (JSON, YAML, XML).
- Markdown: `seldia`, `mermaid` and `plantuml` blocks are drawn by your installed Markdown extensions.
- Fixed: empty table cells showed a stray "b".

## 1.1.0
- Links open safely (web and mail links only).
- Fixed tick boxes; Markdown paragraphs can be justified.

## 1.0.0
- Rich text editing for Word and Markdown, formula engine for spreadsheets, `.odp` viewing, CSV/TSV editing, Markdown WYSIWYG editor with source view.

## 0.0.2
- PDF: view `.pdf` files, fill text/checkbox/radio/dropdown form fields, add text fields, checkboxes and dropdowns, move/resize/delete fields, add fixed text, rotate and delete pages, with Undo/Redo and Find. pdf.js draws, pdf-lib saves.
- Toolbar above every document: Undo, Redo, Save, Find (Ctrl+F), zoom for documents and slides.
- Word: navigation pane (headings), formatting marks (¶), word count.
- Excel: name box and formula bar, freeze first row, per-column text filter (view only).
- See `doc/plan.md` Phase 7 for the features that are planned next.

## 0.0.1
- First version: sandboxed custom editor for `.xlsx`, `.docx`, `.pptx` (read-only), `.odt`, `.ods`.
