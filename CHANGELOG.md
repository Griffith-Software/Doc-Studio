# Changelog

All notable changes to Doc Studio by Griffith Software. Versions before 1.0.1 were not released publicly (1.0.0 was
withdrawn before anyone downloaded it; its notes are included below).

## 1.0.1 - 2026-10-26

First public release. Licensed under Apache License 2.0 with the Commons Clause (source-available: free to use, modify
and share, not to sell).

### Added

- Open, view, edit and save documents (.docx, .doc, .rtf, .odt), spreadsheets (.xlsx, .xls, .xlsb, .ods, .csv),
  presentations (.pptx, .ppt, .odp), PDFs, Markdown, text, web pages, EPUB books, JSON, XML and pictures, keeping
  each file's own formatting. Markdown has Edit, Split and Preview modes.
- Create new Markdown, .docx, .xlsx, .pptx, PDF, .odt, .ods, .odp, .rtf, text, web page, CSV and JSON files from Home,
  the + tab button or File > New.
- Tools: Convert (one file or many), OCR (scan to text, using the text recognition built into Windows), Compare,
  Merge into one PDF, Split a PDF, Compress, Scan (from a scanner connected to the PC), and Phone (send photos and
  files between a phone and the PC over the home network with a QR code).
- Picture editor: crop, rotate, adjust, filters, draw, text, resize, remove background (on this PC), and saving as
  PNG, JPEG, WebP or BMP.
- Split view, pop-out windows, tabs, search across documents and menus, starred and recent files, light and dark
  themes, keyboard shortcuts.
- Version history (a copy at every save, newest 40 per file), backups before a file is first overwritten (newest
  50), and recovery of unsaved work after a crash or power cut (copied every 5 seconds).
- File associations for 35 document types, an "Open with Doc Studio" right-click entry, Settings > Default apps,
  and per-type file icons.
- AI assistants (optional, off until switched on): Claude Code, Claude Desktop, Gemini CLI and OpenAI Codex can use
  Doc Studio through its MCP server. They create new files only and never delete anything.
- Dictionary lookup (Tools > Dictionary, online).

### Changed (release preparation)

- AI assistants are now off until switched on in Settings > AI assistants, and nothing listens for them while they
  are off.
- Converting a document to PDF or pictures no longer loads pictures the document links to on the internet.
- The Dictionary always shows the word and says it is sent to dictionaryapi.dev before looking it up.
- The Phone window explains which firewall choice to make if Windows asks.
- Doc Studio removes its temporary folders when it closes, and no longer changes files in its install folder while
  running.
- Installer: shows the license, offers a desktop shortcut (unticked), and the Installed apps entry shows the name,
  website and support links. Uninstalling asks whether to keep your Doc Studio data (kept unless you choose to
  delete it) and removes the file-type icons and temporary folders Doc Studio made.
- The installer always installs for the current Windows user, with no administrator prompt (the "Anyone who uses
  this computer" choice is gone: file types and icons are registered per user).
- License: Apache License 2.0 with the Commons Clause instead of MIT, with a NOTICE file. LICENSE, NOTICE and
  THIRD_PARTY_NOTICES.md are installed next to Doc Studio.exe.
