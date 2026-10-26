# Doc Studio vs Google Docs, Sheets and Slides

Doc Studio is an independent project. It is not made by, endorsed by, or affiliated with Google, and Google Docs,
Sheets and Slides are names that belong to Google. They are named here only to compare features.

Research for "offer all of the same features", started 2026-10-14. Sources: Google's Docs editors Help (keyboard shortcuts, downloads, transitions and animations), how-to guides, and the current Docs, Sheets and Slides menus and toolbars. This file tracks what is built and what is left.

Status: **done** = built and tested in Doc Studio · **n/a** = needs Google's cloud, so it has no desktop equivalent. As of 0.6.0 every non-cloud item is built.

## Scope

- Doc Studio is a local, offline desktop app. These cloud features are **n/a**: Share and live co-editing, Gemini, Meet, Extensions/Apps Script/AppSheet, Drive smart chips, Explore (web search), Translate document, and assigning comments to people.
- **Local stand-ins:**
  - Version history: a copy is kept on every save.
  - Comments: stored in the file itself.
  - Voice typing: Windows' Win+H dictation works in every text area.
- Every edit is saved into the document's own format (docx/xlsx/pptx/odt…), keeping the original formatting. Google instead converts files into its own format.

## App shell (all documents) — done

- **Menus:** in-window menu bar per document type (File, Edit, View, Insert, Format, Data/Slide/Arrange/Page, Tools, Help) with Alt+letter access and **Search the menus** (Alt+/). Every toolbar control is also reachable from the menus.
- **Toolbar:** undo, redo, print and zoom at the start, as in Google.
- **File:**
  - New for every type, Open, Open recent, Make a copy, Save, Save as.
  - **Download / convert to**, Convert to…, Convert files… (batch).
  - Rename (click the title), Version history (open or restore), Show in folder, Details, Print, Export PDF.
- **Edit:** cut, copy, paste, paste without formatting, select all, and Find and replace (match case, regular expressions, replace all as one undo step, spreadsheet cells included).
- **View:** zoom, full screen, compact mode (Ctrl+Shift+F) and theme.
- **Tools and Help:**
  - Right-click spelling suggestions, add to dictionary, spelling on/off.
  - Word count (Ctrl+Shift+C).
  - Keyboard shortcuts (Ctrl+/), Preferences (smart quotes, automatic substitutions).
  - Compare documents.

## Google Docs → Word editor (ODT and RTF share the generic tools) — done

Items marked *0.6* were added in version 0.6.0.

| Area | Done |
|---|---|
| Toolbar | Paint format, style, font, size −/+, bold/italic/underline/strike, super/subscript, text colour, highlight, line & paragraph spacing, checklist, bullets, numbering, indent, align, link, picture, table grid picker, comment, page break, clear formatting, ***0.6* Editing/Suggesting switch** (suggestions are saved as Word tracked changes) |
| Insert | Table (grid), horizontal line, special characters and emoji, date, checklist, **footnote**, **comment** (on a selection), **bookmark**, **header/footer**, **page numbers** (header/footer, left/centre/right, "Page X of Y"), **table of contents**, page break, ***0.6* equation** (LaTeX-style input written as real Word math), ***0.6* drawing**, ***0.6* chart** (as a picture), ***0.6* section break** (next page/continuous), ***0.6* dropdown chip** (Word drop-down content control) |
| Format | Text case (UPPER/lower/Title), font size steps, line & paragraph spacing (presets, before/after, custom, keep with next), **table rows/columns** (insert above/below/left/right, delete row/column/table), **page setup** (paper size, orientation, margins, columns, page colour), ***0.6* list styles** (bullets, arrows, checks, squares, dashes, stars, 1. 1) a. A. i. I.; restart numbering), ***0.6* image options** (size, crop, alt text; drag handle to resize), ***0.6* borders & shading**, ***0.6* update style to match** |
| View | Outline panel, paragraph marks, zoom, ***0.6* ruler with draggable left/first-line/right indents**, ***0.6* pageless mode** |
| Tools | Word count, comments list (resolve/delete), compare documents, preferences, spelling, ***0.6* review suggested edits** (accept/reject one or all), ***0.6* citations** (APA/MLA/Chicago, bibliography), ***0.6* dictionary** (needs internet) |

Shortcuts as in Google: Ctrl+Shift+./, (font size), Ctrl+Shift+9 (checklist), Ctrl+Enter (page break), Ctrl+Alt+F (footnote), Ctrl+Alt+M (comment), Ctrl+\ (clear formatting), Ctrl+./, (super/subscript), Alt+Shift+5 (strikethrough).

## Google Sheets → workbook editor — done

| Area | Done |
|---|---|
| Toolbar | Paint format, zoom, currency, percent, decimals, number formats, font, size −/+, bold/italic/underline/strike, colours, borders, merge, horizontal/vertical align, wrap, **rotation**, link, note, chart, functions (Σ menu with more functions), filter, sort, freeze, conditional formatting |
| Edit | Find & replace in cells, paste values, delete values, ***0.6* paste special** (values, format, formulas, transposed, column widths; Ctrl+Shift+V), ***0.6* delete cells** (shift left/up; references follow, deleted ones become #REF!) |
| View | Gridlines, show formulas, zoom, freeze panes |
| Insert | **Chart** (column, bar, line, area, pie, scatter, written as real Excel charts), **pivot table** (a summary sheet with live SUMIFS/COUNTIFS/AVERAGEIFS/MINIFS/MAXIFS formulas), **image over cells**, **function picker** (every function, searchable), **link** (HYPERLINK, Ctrl+click opens), **checkbox**, **dropdown**, **note** (Excel comments with marker and hover text), rows/columns, new sheet, ***0.6* cells** (shift right/down), ***0.6* drawing** |
| Format | Custom number format, vertical align, rotation, **alternating colours**, **convert to table**, conditional formatting, clear formatting, ***0.6* theme** (colours and fonts) |
| Data | Sort range (with header row, by column), filter, **remove duplicates**, **trim whitespace**, **split text to columns**, **randomize range**, **column stats**, **named ranges**, **protect sheet**, **data validation** (list, number range, checkbox), ***0.6* filter views** (save, apply, delete), ***0.6* slicers**, ***0.6* protected ranges**, ***0.6* group/ungroup rows and columns** with collapse/expand (Excel outline; Alt+Shift+→/←) |
| Sheet tabs | Rename, **duplicate, delete (formulas become #REF!), hide/unhide, colour, move left/right**, add, ***0.6* copy to another open workbook** (Tools menu) |

## Google Slides → presentation editor — done

| Area | Done |
|---|---|
| Toolbar | **Present**, new slide by layout, duplicate/delete slide, text box, shape, picture, line/arrow, table, fill, border, arrange, background, layout, theme, transition, animate, font, size, bold/italic/underline/strike, colour, align |
| View | **Slideshow** (full screen, transitions, click animations, keyboard, black screen, go to slide), **presenter view** (timer, next slide, speaker notes), **grid view**, zoom, **smart guides** (snap to slide and other shapes) |
| Insert | Table (grid, cell text editable), line, arrow, word art, special characters, link (on text or shape), **slide numbers**, **animation** (appear, fade in, fly in), ***0.6* video and audio** (PowerPoint media; play in the editor and the slideshow), ***0.6* chart** (a real chart with its data in an embedded workbook; Format > Edit chart data), ***0.6* diagram** (process, timeline, cycle, hierarchy as grouped shapes in theme colours), ***0.6* drawing**, ***0.6* comments** (PowerPoint comments, pinned on the slide; Ctrl+Alt+M) |
| Format | **Fill colour, border colour/weight/dash, drop shadow, size & rotation**, ***0.6* crop**, ***0.6* image options** (brightness, contrast, transparency, greyscale, reflection) |
| Slide | New, duplicate, delete, move, **skip slide, background colour (one or all), apply layout, transition (fade, push, wipe, cover, split, zoom, dissolve, cut; speed; all slides), change theme (colours and fonts)**, ***0.6* theme builder** (custom colours, heading/body fonts, master background, title/body text sizes) |
| Arrange | **Order (front, forward, backward, back), align (left/centre/right/top/middle/bottom; to slide or to each other), distribute, centre on page, rotate 90°, flip, group, ungroup** |

### Known limits

- **Suggesting mode:** records typed and deleted text. Paragraph breaks and formatting changes made while suggesting are applied directly.
- **Citation sources:** kept on this computer for each document, not inside the file.
- **Filter views, slicers and protected ranges:** saved inside the workbook (a custom document property). Excel ignores them; Doc Studio applies them again when the file is reopened.
- **Theme builder:** changes the theme and the slide master. Placeholder layouts are not redrawn shape by shape.

## Conversions — done

Use **File > Download / convert to**, **Convert to…** (dialog), or **Convert files…** (batch, saved next to each original). With Word, Excel, PowerPoint or LibreOffice installed, office-to-office and office-to-PDF conversions go through them; every route also works without them.

| From | To |
|---|---|
| Markdown | Word, PDF, ODT, RTF, HTML, TXT, EPUB |
| Word (.docx) | PDF, ODT, RTF, DOC*, Markdown, HTML, TXT, EPUB |
| ODT / RTF | Word, PDF, the other of ODT/RTF, DOC*, Markdown, HTML, TXT, EPUB |
| HTML / text | PDF, Word, ODT, RTF, Markdown, (HTML), EPUB |
| Excel / CSV / older sheets | PDF, ODS, XLS, XLSB, XLSX, CSV, TSV, HTML, Markdown table, JSON |
| PowerPoint | PDF (one page per slide), PNG (one per slide), ODP*, PPT*, Markdown/TXT outline with speaker notes |
| PDF | Word, Markdown, TXT, HTML, PNG per page |
| Images | PDF, PNG, JPEG, WebP |

\* needs Office or LibreOffice.

## Tests

- `tests/shell-test.js`: 34 checks (menus, menu search, find & replace, versions, zoom, every conversion).
- `tests/docs-tools-test.js`: 25 checks (Word features).
- `tests/sheets-tools-test.js`: 25 checks.
- `tests/slides-tools-test.js`: 25 checks.
- `tests/docs-extra-test.js`: 19 checks (suggesting, review, equations, lists, borders, ruler, section breaks, pictures, dropdowns, styles, charts, pageless).
- `tests/sheets-extra-test.js`: 19 checks (paste special, insert/delete cells, groups, filter views, slicers, protected ranges, theme, drawing, copy sheet).
- `tests/slides-extra-test.js`: 18 checks (image options, video, audio, charts, diagrams, comments, theme builder, slideshow media).

Outputs were also opened in real Word, Excel and PowerPoint (`tests/office-validate.ps1` plus checks over COM).
