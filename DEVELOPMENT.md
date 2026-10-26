# Doc Studio: developer notes

How Doc Studio is built, how each editor works and how to run the tests. For what the program does and how to install it, see README.md.


Windows desktop app (Electron) for viewing, editing, creating and converting documents. It keeps each file's own formatting. Copyright (c) 2026 Richard Griffith (d/b/a Griffith Software). Licensed under the Apache License 2.0 with the Commons Clause (see LICENSE and NOTICE). Settings, version history and backups saved under the program's earlier name are carried over on first start.

## The window (v0.7.0)

Built from the Doc Studio mockup (renderer/index.html, renderer/ui-shell.js, renderer/shell.css), light and dark:

- **App bar** (the title bar; Windows draws minimise/maximise/close over it): logo, **Search documents…** (open tabs, recent, starred and files in Documents, Desktop, Downloads and cloud folders, plus every menu command; Alt+/), light/dark switch, **Settings** (appearance, your name for comments, thumbnails, page layout, status bar, spelling).
- **Action bar:** **Open** (▾ new, recent), **Save**, **Print**, **Edit** (switches View/Edit), **Annotate** (highlighting for PDFs; comments, notes and suggestions for the others), **Convert ▾** (this document to any format, batch conversion, export as PDF), **More ▾** (every menu of the current document: File, Edit, View, Insert, Format, …; Alt+letter opens them here). While editing, the menus also show as a menu bar above the editing toolbar.
- **Sidebar:** Home (start page with new documents, recent, starred and tools), Documents (browse Documents, Desktop, Downloads, Home or any folder), Recent, Starred, Cloud Storage (OneDrive, Google Drive, Dropbox, iCloud and Box folders synced to this PC), then the tools: **Convert**, **OCR (Scan to Text)** (Windows' built-in text recognition, offline; pictures and PDFs to a Word document, text, Markdown or the clipboard), **Compare**, **Merge** (PDFs, pictures and any document into one PDF), **Split** (every page, every N pages, ranges, or extract), **Compress** (PDFs losslessly or as pictures, pictures, and Office/ODF files by shrinking their pictures) — renderer/tools-panel.js. Narrow windows show icons only; Ctrl+Shift+F hides the sidebar and menus.
- **Tabs** with a **+** for a new document of any type.
- **Page bar and toolbar (two lines, renderer/toolbar-layout.js):**
  - **Line 1:** previous/next page, page number / count, zoom −/+ and size, then the main editing tools for the document (font, size, bold, colours, alignment...), then View/Edit/Split, star, and single page / two pages / thumbnails.
  - **Line 2:** the menus (File, Edit, View... while editing) and the secondary tools. Tools that don't fit go into a » menu rather than a third line.
  - **Toolbar settings (the gear at the end of line 2):** choose for each tool Line 1, Line 2 or Menus only, and drag or use ↑ ↓ to reorder; remembered per kind of document, with Reset to default.
- **Window buttons** (split view, pop out, full screen) sit at the right end of the action bar. The **sidebar button** at the top left shows or hides the sidebar (remembered); compact mode (Ctrl+Shift+F) hides the sidebar and menus, and the menus then offer to show them again.
- **Several documents at once** (renderer/popout.js): pop a document into its own window with the ⧉ button on the page bar, by right-clicking its tab, from View, or by dragging its tab out of the window. **Open side by side** puts both windows on the two halves of the screen (read instructions in one while writing in the other); **Arrange all windows side by side** tiles every Doc Studio window. Drag a tab onto another Doc Studio window to move it there, or along the tab bar to reorder. Unsaved edits travel with the document, and it only leaves the first window once the new one has it open. Each window has its own tabs, menus and dialogs.
- **Split view** (renderer/split.js): two documents side by side in one window, using the split button on the page bar, a tab's right-click menu, or View › Split view. Each half has a header (name, swap ⇄, close ✕); click in a half to work in it, and the page bar, toolbars and menus follow. A tab chosen from the tab bar goes into the half you are working in. Drag the divider to resize it (double-click resets it), and an empty half offers open documents, Open a file… or a new document.
- **Thumbnails** on the right for Word, PowerPoint and PDF pages (PDFs keep their own: drag to reorder and right-click page actions while editing). PDF **View** now uses the same page renderer as Edit, read-only, so page numbers, thumbnails, zoom and printing work the same.
- **Default apps** (v0.7.1–0.7.2): the installer (build/installer.nsh) registers Doc Studio for 35 document types: it appears in File Explorer's Open with list, in a right-click **Open with Doc Studio** entry (Windows 11: Show more options) and in Windows' Default apps. Windows lets only the owner choose a default, and only File Explorer's Open with offers **Always**, so Settings › **Make Doc Studio the default…** (also More › Default apps…) lists every type with the app that opens it now, then shows a sample file for each ticked type in File Explorer with the three clicks to make; opening the sample in Doc Studio marks that type done and moves on.
- **Markdown first**: Markdown leads Start something new, Open ▾, File › New and the tab + menu (AI tools write Markdown; it is what sets Doc Studio apart).
- **Recovery** (renderer/recovery.js): every 5 seconds, each document with new unsaved edits is copied to <userData>/recovery (written to a temporary name, then renamed, so a power cut never leaves half a copy). The document's own file is untouched. Saving, closing or Don't save removes the copy. If Doc Studio didn't close normally, the next start offers **Recover unsaved work** (Restore all / Discard). A slide text box or PDF text box being typed in is copied once you leave it.
- **Version** at the right end of the status bar.
- **Generic names** (v0.7.4): Doc Studio's own features say Document (.docx), Spreadsheet (.xlsx), Presentation (.pptx), Slide show, Text art, never Word/Excel/PowerPoint/WordArt (Microsoft trademarks). Those names appear only where Doc Studio really means Microsoft's programs ("Microsoft Word and LibreOffice are not installed", "Opened for editing with Microsoft Excel") or says it opens their files. File types registered by the installer read "DOCX Document", "XLSX Spreadsheet", "PPTX Presentation"…
- **Message boxes** (preload.js messageBox): questions (Save changes?, Discard?) and errors are drawn in the page in Doc Studio's style instead of Windows message boxes; desk.confirm/desk.error keep Electron's showMessageBox options (buttons, defaultId, cancelId, type). Enter answers with the focused button, Esc with the cancel button.
- **Tool windows** (v0.7.5, renderer/tools-panel.js): Convert, OCR, Compare, Merge, Split and Compress all open the same kind of window: what the tool does, the file(s) it works on (a row with Choose file… / Use the open document), only the options that apply, and one button that stays off until the tool has what it needs. Convert takes several files and offers only the formats they can all become (saved next to each original or in a chosen folder); Split shows how many files the settings make and where they go; results never overwrite an existing file ("name (2).pdf"). Convert to… / Convert files… and the editors' Compare documents… open the same windows.
- **Scan** (v0.8.0, tools-panel.js scan, main.js scan:*): scans from a scanner connected to this PC through Windows Image Acquisition (WIA, built into Windows; nothing to install). Choose the scanner, colour/greyscale/black-and-white, 150-600 dpi and Letter/A4/whole glass, scan one or more pages, then make a picture to edit (each page), a PDF (JPEG pages, or PNG for black-and-white text), or text by OCR (document or Markdown). Some drivers return BMP whatever is asked; that is read as it is.
- **Image editor** (v0.8.0, renderer/image-editor.js): pictures open with View and Edit. Edit has Crop (free, original, square, 4:3, 3:4, 16:9, 9:16; rotate, flip, straighten), Adjust (brightness, contrast, saturation, warmth, tint, highlights, shadows, sharpness, vignette), Filters (vivid, warm, cool, mono, noir, sepia, fade, dramatic, vintage, with strength), Draw (pen, highlighter), Text, Background and Resize, with Undo/Redo/Revert on the toolbar. Light, colour and filters stay adjustable settings on top of the picture until saving. Save writes PNG, JPEG, WebP or BMP (Save As changes the type; BMP is written by Doc Studio itself), and Convert turns pictures into those or PDF.
- **Remove background** (renderer/bg-remove.js): the U2-Net-p model (renderer/vendor/bgremove/u2netp.onnx, Apache-2.0, from rembg) run by ONNX Runtime Web (renderer/vendor/ort, MIT; copied from the onnxruntime-web dev dependency) on this PC, loaded on first use. Then Transparent/White/any colour behind the subject, and an Erase more / Bring back brush. Notices: renderer/vendor/THIRD-PARTY-NOTICES-bgremove.txt. The page allows wasm-unsafe-eval for it (Content-Security-Policy).
- **Phone** (v0.9.0, phone.js + phone-page.html, tools-panel.js phone): send photos and files both ways between a phone and this PC over the home network, with no app on the phone. The window shows a QR code (qrcode-generator, MIT, renderer/vendor/qrcode); the phone's camera opens a page served by Doc Studio with Take a photo / Choose files and the files the PC offers. Received files are listed (Open, Save all to a folder…, Open all); Send to your phone takes the open document (as it is now) or files. Safety: a random 32-character code in every address (403 without it), only private network addresses answered, uploads go to a temporary folder (deleted when the window closes) with safe names and a 200 MB limit, the phone can download only what was offered, and the link stops when the window closes or after 15 minutes unused. Windows Firewall asks once to allow Doc Studio on private networks. Plain HTTP on the home network (no certificate for a home PC).
- **AI assistants** (v0.10.0): Claude Code, Claude Desktop, Gemini CLI and OpenAI Codex can use Doc Studio through one MCP server (integrations/mcp/mcp-server.js, no dependencies; 15 docstudio_* tools: status, list/read open documents, read any document as Markdown, create .md/.txt/.html/.docx/.pdf from Markdown, rewrite, find/replace, copy, convert, merge, split, compress, OCR, edit pictures, scan, show). It talks to the running app through ai/bridge.js (127.0.0.1, random port, secret key in <userData>/ai-bridge.json; Doc Studio is started when needed), which runs the work in a hidden worker window (renderer/ai-worker.js, index.html?worker=1; its files stay out of Recent) and enforces the owner's rules: new files only (free names when no output is given), copy first to change someone else's file, files an AI made may change (same session freely, an earlier session asks Allow once / Allow for this file / Don't allow, recorded in <userData>/ai-files.json), never delete or upload; Settings > AI assistants has an off switch. tools/make-ai-packages.js (run before every build) puts the packages in resourcesai: a Claude Code plugin with its own marketplace, doc-studio.mcpb for Claude Desktop (runs on Claude's Node), a Gemini CLI extension, and mcp-server.js for Codex (run with Doc Studio.exe and ELECTRON_RUN_AS_NODE=1, so no Node install is needed). Settings > AI assistants shows the commands with this PC's paths. ChatGPT's app only reaches internet tools, so it is not supported. tests/ai-test.js (22).
- **File icons** (v0.11.0): files Doc Studio opens show their type in File Explorer: a page with the Doc Studio logo and a coloured band (DOCX blue, XLSX/CSV green, PPTX orange, PDF red, MD dark grey...; generic colours and the extension, no other company's logos). Icons: build/filetypes/<ext>.ico from tools/make-filetype-icons.py (Pillow), shipped in resourcesiletypes and set on the DocStudio.<ext> types. Windows 11 records any choice of Doc Studio as the whole app (one icon, tested: per-type types, hard-linked program names and clearing the Open with lists all still record ApplicationsDoc Studio.exe), so a small File Explorer icon handler (iconhandler/IconHandler.cs, built by tools/make-icon-handler.js with the C# compiler that comes with Windows) answers with the icon for each file's extension. shell-icons.js copies it to %LOCALAPPDATA%Doc Studio Icons<version> (Explorer keeps it loaded, so it is never overwritten in place) and registers it per user at every start; the uninstaller removes it. File Explorer shows the new icons after it restarts.

## What it handles

| | Formats | How |
|---|---|---|
| View + edit | `.docx` | **View** = paginated layout. **Edit** = the same rendering, editable (body, headers, footers, footnotes, endnotes), saved by patching the original XML. |
| View + edit | `.xlsx .xlsm` (`.xltx/.xltm` save as workbooks) | Spreadsheet editor with live formulas that keeps the file's own formatting. |
| View + edit | `.pptx` (`.pptm/.ppsx` save as `.pptx`) | Own slide renderer (View) and slide editor (Edit). |
| View + edit | PDF | **View** = pdf.js pages, read-only. **Edit** = the same pages with the pdf-lib editor (text, annotations, forms, signatures, pages). |
| View + edit | `.odt` | Own OpenDocument editor, patched in place. |
| View + edit | `.rtf` | Through Word when installed (full fidelity), otherwise Doc Studio's own RTF editor. |
| View + edit, through Office/LibreOffice | `.doc .xls .xlsb .ods .ppt .odp` | Converted to the modern format with the installed office program, edited in the full editor, then converted back on save. Without Office, spreadsheets fall back to SheetJS (values, formulas and number formats only) and `.doc` opens as text only. |
| View + edit | `.md`, `.html`, `.csv/.tsv`, text and code | Markdown Edit/Split/Preview (Split keeps both panes scrolled to the same block, whichever side is scrolled); plain-text editor that keeps encoding and line endings; HTML preview; CSV table view. |
| View | `.epub`, `.numbers` (edit via SheetJS, Save As `.xlsx`), images, audio, video | Own EPUB reader; SheetJS; native |

**New** (File > New, or the welcome page): Word, Excel, PowerPoint, PDF, OpenDocument text/spreadsheet/presentation, RTF, Markdown, text, HTML, CSV, JSON. **Save As** offers every format the current editor can write on this PC (for example a Word document as `.odt`, `.rtf`, `.doc` or `.pdf`; a workbook as `.ods`, `.xls`, `.xlsb` or `.csv`; a presentation as `.odp`, `.ppt` or `.pdf`).

Also: tabs, drag and drop, recent files, print, export to PDF, light/dark/system theme, single instance, unsaved-changes prompts, and a backup of every file before Doc Studio first overwrites it (`%APPDATA%\Doc Studio\backups`, File > Open Backups Folder, newest 50 kept).

**Google Docs/Sheets/Slides-style features** (all non-cloud features as of v0.6.0; see [FEATURE-PARITY.md](FEATURE-PARITY.md) for the full list):
- **Every document:**
  - Menus per document type (under More, and as a menu bar while editing), with search (Alt+/).
  - Undo/redo on the editing toolbar; print, zoom and pages on the action and page bars.
  - Find and replace (match case, regex).
  - Version history on every save (`%APPDATA%\Doc Studio\versions`, newest 40 per file), Details, Rename, Make a copy.
  - Spelling suggestions on right-click, word count, keyboard-shortcut sheet, compact mode.
  - **Download / convert to** any format and **Convert files…** for batches (renderer/convert-ui.js).
  - A **+** button after the open tabs creates a new document of any type.
- **Word:** footnotes, comments, bookmarks, headers/footers, page numbers, table of contents, horizontal line, special characters, checklists, paint format, line & paragraph spacing, page setup (size, orientation, margins, columns, page colour), table rows/columns, change case, outline, paragraph marks, compare documents, smart quotes and substitutions (renderer/docx-tools.js, renderer/docs-common.js). Also suggesting mode saved as tracked changes with a review panel, equations (Word math), section breaks, dropdown chips, list styles, picture size/crop/alt text with a resize handle, borders & shading, update style to match, a ruler with draggable indents, pageless view, citations, dictionary, drawings and charts (renderer/docs-extra.js).
- **Sheets:**
  - Charts, pivot tables (live formulas), pictures, notes, checkboxes, dropdowns, links, function picker.
  - Remove duplicates, trim, split to columns, randomize, column stats, named ranges, protect sheet, data validation.
  - Alternating colours, convert to table, rotation, vertical align, gridlines, show formulas, zoom.
  - Sheet tabs: duplicate, delete, hide, colour, move (renderer/sheet-tools.js).
  - Paste special, insert/delete cells with shifting, group/collapse rows and columns, filter views, slicers, protected ranges, workbook theme, drawings, copy a sheet to another open workbook (renderer/sheet-tools2.js).
- **Slides:**
  - Slideshow with transitions and click animations, presenter view, grid view, smart guides.
  - Arrange (order, align, distribute, rotate/flip, group).
  - Fill/border/shadow/size, backgrounds, layouts, themes, transitions, animations, tables, lines/arrows, word art, slide numbers, links, skip slide (renderer/pptx-tools.js).
  - Crop and image options (brightness, contrast, transparency, greyscale, reflection), video and audio, real charts with an embedded workbook, diagrams, comments, theme builder (renderer/pptx-tools2.js).

## How each editor keeps the original formatting

### Word (renderer/docx-editor.js)

Edit mode renders a *marked copy* of each part (document, headers, footers, footnotes, endnotes) with docx-preview: every paragraph gets a bookmark and every run an invisible private-use tag, which become `data-dd-p` / `data-dd-r` attributes. On save, unchanged paragraphs are kept byte-for-byte; edited ones are rebuilt from their own `pPr`/`rPr` plus only the formatting the user changed. Pictures, fields, footnote references and markers are copied as-is. Lists, tables, pictures and page breaks are applied to the XML and re-rendered. Pasted content keeps bold/italic/underline/colour/links (Ctrl+Shift+V pastes plain text). Undo reaches back across structural edits. View splits the content into real pages.

### Excel (renderer/sheet-*.js)

- **Model** saves by patching only what changed (cells, appended styles, widths, merges, sheets). Charts, pictures, macros, pivot tables and every other part are kept. When formulas change, `calcChain.xml` is dropped and `fullCalcOnLoad` set.
- **Formulas**: fast-formula-parser + @formulajs/formulajs (MIT, bundled into `renderer/vendor/formula.js`), plus Doc Studio's XLOOKUP, MATCH, INDIRECT, OFFSET, FILTER, SEQUENCE, Excel-accurate IF, table (structured) references, dynamic-array spilling and external links (from the values cached in the file). Dependency-graph recalculation; circular references show 0 with a warning.
- **Grid**: virtualised; real widths/heights, fonts, fills, borders, number formats, merges, frozen panes, conditional formatting (colours, colour scales, data bars, icon sets), table styles, charts (Chart.js) and pictures.
- **Editing**: everything a spreadsheet user expects (typing, references by clicking, F4, copy/paste with formulas, fill, undo/redo, formatting, AutoSum, merge), plus insert/delete/hide rows and columns (formulas, names, merges, tables, charts, conditional formatting and validation all shift), sort, AutoFilter (table-aware), conditional formatting rules, freeze panes.

### PowerPoint (renderer/pptx-editor.js, pptx-ui.js)

Slides are drawn from their own XML with the full inheritance chain (slide, layout, master, theme): placeholders, text styles, theme colours and fonts, backgrounds, preset and custom shapes, pictures, tables, groups, connectors, charts and SmartArt drawings. Editing: double-click text to edit (runs keep their own properties; new formatting is added to them), move/resize/nudge/delete shapes, insert text boxes, shapes and pictures, new slide from any layout, duplicate, delete, drag to reorder, speaker notes, undo/redo. Saving rewrites only changed parts and keeps relationships and content types consistent.

### PDF (renderer/pdf-editor.js)

pdf.js (legacy build, main thread) renders pages with a text layer; pdf-lib writes changes. Tools: add text (Windows fonts embedded), edit existing text (covers the old line and writes the new one in the same font and size), highlight (text or area), freehand ink, rectangle, ellipse, line, white-out, sticky notes, pictures and drawn signatures, form filling. Pages: rotate, delete, insert blank, insert from another PDF, drag to reorder. Annotations get their own appearance streams so every viewer shows them.

### OpenDocument text (renderer/odt-editor.js)

`content.xml` is drawn with the document's own styles (style chains, default styles, font faces, list styles, tables, pictures, page size, header/footer). Saving rebuilds only changed paragraphs: their runs reuse the original span styles and new formatting goes into new automatic styles. New paragraphs copy the one they came from; deleted ones are removed. Fields, notes, bookmarks, frames and annotations are carried over untouched; `styles.xml`, meta and settings are written back unchanged.

### Rich Text (renderer/rtf-editor.js)

With Word installed, RTF opens through Word (converted to `.docx` for editing, converted back on save). Otherwise rtf.js draws the file and Doc Studio writes RTF from the page: fonts, sizes, colours, bold/italic/underline/strike, super/subscript, highlight, alignment, indents, spacing, line breaks and pictures, with the original page size and margins. An unchanged file is saved byte for byte.

### Conversions (convert.js)

Word, Excel and PowerPoint are driven over COM by a small PowerShell script (files opened read-only with macros disabled, one conversion at a time with one retry); LibreOffice is run headless when installed. Conversions go through a private temp folder.

## Tests

All tests run the real app (Electron) on real files; paths go in the environment because the app opens any file named on its command line. Use Node.js 22 or later; `electron` is `node_modules\.bin\electron` after `npm install`.

| Test | Checks |
|---|---|
| `DOCSTUDIO_SRC=<rich.docx> DOCSTUDIO_OUT=<dir> electron tests/docx-fidelity.js` | 24: no-edit identity, text/format edits, links, splitting, structural edits, new documents |
| `DOCSTUDIO_SRC=<word-made.docx> … tests/docx-parts-test.js` | 15: headers/footers/footnotes, formatted paste, undo, pagination |
| `DOCSTUDIO_SRC=<budget.xlsx> … tests/xlsx-test.js` | 46: values, editing, formatting, undo, clipboard, saving |
| `DOCSTUDIO_OUT=<dir> electron tests/xlsx-ui-test.js` | 16: real keyboard and mouse input |
| `DOCSTUDIO_SRC=<excel-made.xlsx> … tests/xlsx-features-test.js` | 33: tables, charts, CF, validation, external links, spills, rows/columns, sort, filter |
| `DOCSTUDIO_SRC=<report.pdf> DOCSTUDIO_FORM=<form.pdf> … tests/pdf-test.js` | 18: every PDF tool, forms, page operations |
| `DOCSTUDIO_SRC=<deck.pptx> … tests/pptx-test.js` | 17: text, shapes, notes, inserts, undo, slide operations, new deck |
| `DOCSTUDIO_ODT=<file.odt> DOCSTUDIO_RTF=<file.rtf> … tests/odt-test.js` | 15: ODT edits/styles/pictures/new file, native RTF writer |
| `DOCSTUDIO_SRC=<folder> … tests/convert-test.js` | 10: `.doc .xls .ods .ppt .odp` through Office, Save As formats |
| `DOCSTUDIO_SRC=<folder> … tests/shell-test.js` | 34: menus, menu search, find & replace, versions, zoom, every conversion route |
| `DOCSTUDIO_SRC=<rich.docx> … tests/docs-tools-test.js` | 25: footnotes, comments, bookmarks, page numbers, headers/footers, TOC, spacing, page setup, tables, checklists |
| `DOCSTUDIO_SRC=<budget.xlsx> … tests/sheets-tools-test.js` | 25: charts, pictures, notes, checkboxes, dropdowns, data cleanup, named ranges, pivot, tables, tabs, protection |
| `DOCSTUDIO_SRC=<deck.pptx> … tests/slides-tools-test.js` | 25: arrange, group, format, animations, transitions, backgrounds, themes, layouts, tables, slideshow, grid view |
| `DOCSTUDIO_SRC=<long.md> … tests/md-scroll-test.js` | 6: Markdown Split view scroll sync |
| `DOCSTUDIO_OUT=<dir> electron tests/toolbar-test.js` | 24: two toolbar lines, overflow » menu, toolbar settings (place, reorder, menus only, reset), window buttons, sidebar button, compact labels |
| `DOCSTUDIO_MD=<file.md> DOCSTUDIO_DOCX=<file.docx> DOCSTUDIO_XLSX=<file.xlsx> … tests/split-test.js` | 22: split view, active half, editing in both halves, replace, swap, resize, Home, close, picker |
| `DOCSTUDIO_MD=<file.md> DOCSTUDIO_DOCX=<file.docx> … tests/popout-test.js` | 19: pop out with unsaved edits, side by side, save from the new window, move between windows, drag out, arrange windows |
| `DOCSTUDIO_DOCX=<multi-page.docx> DOCSTUDIO_PDF=<report.pdf> … tests/shell-ui-test.js` | 37: window, sidebar pages, page bar, thumbnails, layouts, starring, search, menus under More, Edit/Annotate, PDF view, merge, split, compress, OCR, light/dark |
| `DOCSTUDIO_SRC=<rich.docx> … tests/docs-extra-test.js` | 19: suggesting and review, equations, lists, borders, ruler, section breaks, pictures, dropdowns, styles, charts |
| `DOCSTUDIO_SRC=<budget.xlsx> … tests/sheets-extra-test.js` | 19: paste special, insert/delete cells, groups, filter views, slicers, protected ranges, theme, drawing, copy sheet |
| `DOCSTUDIO_SRC=<deck.pptx> DOCSTUDIO_VIDEO=<clip.mp4> … tests/slides-extra-test.js` | 18: image options, video, audio, charts, diagrams, comments, theme builder |
| `npm test` | the syntax check, then image-test, toolbar-test and xlsx-ui-test (no sample files needed) |
| `electron tests/smoke.js <samplesDir> <outDir>` | smoke test over a folder of samples |
| `tests/office-validate.ps1 <files…>` | opens outputs in real Word/Excel/PowerPoint and reports what they see |

## Commands

- `npm start`: run from source
- `npm run bundle`: rebuild the formula engine and PDF bundles and their third-party notices
- `npm run dist`: build `dist/DocStudio-Setup-<version>.exe` (NSIS, per-user, unsigned; license page, optional desktop shortcut, keep/delete data on uninstall: build/installer.nsh; OriginalFilename: packaging/afterSign.js)
- `node tests/check-package.js`: after a build, confirms every renderer script, the converter and the templates are inside app.asar and that build-only packages are not

## Notes

- The renderer is sandboxed (contextIsolation, no Node, CSP without eval). Node-only work (blank `.docx`, `.doc` text, backups, Windows fonts for PDFs, conversions) runs in the main process over IPC.
- Links in documents open in the default browser. HTML previews and EPUB chapters run without scripts.
- Dependencies are permissively licensed (MIT/Apache/BSD); no GPL/AGPL code.
- Untested: LibreOffice conversions (no LibreOffice on the build PC), very large presentations.
