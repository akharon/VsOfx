# Changelog

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
