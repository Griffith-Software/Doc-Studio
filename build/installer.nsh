; Doc Studio file registration (per user, HKCU).
; - Adds Doc Studio to "Open with" for every document type it reads (OpenWithProgids), each with its own icon
;   (resources\filetypes\<ext>.ico: a page with the Doc Studio logo and a coloured band naming the type). Doc Studio
;   is offered by type, not as one app, so choosing it with Always records the type, and Explorer shows its icon.
; - Adds "Open with Doc Studio" to the right-click menu (Windows 11: under "Show more options").
; - Registers Doc Studio with Windows Default apps (Capabilities + RegisteredApplications), so the owner can pick it
;   as the default there. Windows does not let an installer set the default itself.

; - The finish page offers a desktop shortcut (unticked); uninstalling asks whether to keep the user's data (kept
;   unless they choose to delete it; a silent uninstall keeps it, or deletes it with --delete-app-data).

!include "WordFunc.nsh"

!define DS_EXE "$INSTDIR\Doc Studio.exe"
!define DS_CAPS "Software\Griffith Software\Doc Studio\Capabilities"

; Always install for the current Windows user (no "Anyone who uses this computer" page, no administrator prompt):
; file types, "Open with" entries and the icon handler are registered per user, so an all-users install would
; only work fully for the user who installed it.
!macro customInstallMode
  StrCpy $isForceCurrentInstall "1"
!macroend

; The installer's own version resource: electron-builder leaves these two empty.
!macro customHeader
  VIAddVersionKey /LANG=1033 OriginalFilename "DocStudio-Setup-${VERSION}.exe"
  VIAddVersionKey /LANG=1033 InternalName "DocStudio-Setup"
!macroend

; Finish page: run Doc Studio (ticked) and create a desktop shortcut (unticked).
!macro customFinishPage
  Function DSStartApp
    ${if} ${isUpdated}
      StrCpy $1 "--updated"
    ${else}
      StrCpy $1 ""
    ${endif}
    ${StdUtils.ExecShellAsUser} $0 "$launchLink" "open" "$1"
  FunctionEnd
  Function DSDesktopShortcut
    CreateShortCut "$DESKTOP\${SHORTCUT_NAME}.lnk" "$appExe" "" "$appExe" 0 "" "" "${APP_DESCRIPTION}"
    ClearErrors
    WinShell::SetLnkAUMI "$DESKTOP\${SHORTCUT_NAME}.lnk" "${APP_ID}"
    System::Call 'Shell32::SHChangeNotify(i 0x8000000, i 0, i 0, i 0)'
  FunctionEnd
  !define MUI_FINISHPAGE_RUN
  !define MUI_FINISHPAGE_RUN_FUNCTION "DSStartApp"
  !define MUI_FINISHPAGE_SHOWREADME
  !define MUI_FINISHPAGE_SHOWREADME_TEXT "Create a desktop shortcut"
  !define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED
  !define MUI_FINISHPAGE_SHOWREADME_FUNCTION "DSDesktopShortcut"
  !insertmacro MUI_PAGE_FINISH
!macroend

!macro DSAssoc EXT DESC
  WriteRegStr HKCU "Software\Classes\DocStudio.${EXT}" "" "${DESC}"
  WriteRegStr HKCU "Software\Classes\DocStudio.${EXT}\DefaultIcon" "" '"$INSTDIR\resources\filetypes\${EXT}.ico"'
  WriteRegStr HKCU "Software\Classes\DocStudio.${EXT}\shell\open\command" "" '"${DS_EXE}" "%1"'
  WriteRegStr HKCU "Software\Classes\.${EXT}\OpenWithProgids" "DocStudio.${EXT}" ""
  WriteRegStr HKCU "${DS_CAPS}\FileAssociations" ".${EXT}" "DocStudio.${EXT}"
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\shell\DocStudio" "" "Open with Doc Studio"
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\shell\DocStudio" "Icon" '"${DS_EXE}",0'
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\shell\DocStudio\command" "" '"${DS_EXE}" "%1"'
!macroend

!macro DSUnassoc EXT
  ; The icon Doc Studio gave Windows' own entry for the type (shell-icons.js fixAutoFiles), if it is still ours.
  ReadRegStr $0 HKCU "Software\Classes\${EXT}_auto_file\DefaultIcon" ""
  StrCmp $0 '"$INSTDIR\resources\filetypes\${EXT}.ico"' 0 +2
    DeleteRegKey HKCU "Software\Classes\${EXT}_auto_file\DefaultIcon"
  DeleteRegKey HKCU "Software\Classes\DocStudio.${EXT}"
  DeleteRegValue HKCU "Software\Classes\.${EXT}\OpenWithProgids" "DocStudio.${EXT}"
  DeleteRegKey HKCU "Software\Classes\SystemFileAssociations\.${EXT}\shell\DocStudio"
!macroend

; One list for install and uninstall.
!macro DSTypes M
  !insertmacro ${M} "docx" "DOCX Document"
  !insertmacro ${M} "docm" "DOCM Document"
  !insertmacro ${M} "dotx" "DOTX Template"
  !insertmacro ${M} "doc" "DOC Document"
  !insertmacro ${M} "dot" "DOT Template"
  !insertmacro ${M} "rtf" "Rich Text document"
  !insertmacro ${M} "odt" "OpenDocument text"
  !insertmacro ${M} "ott" "OpenDocument text template"
  !insertmacro ${M} "md" "Markdown document"
  !insertmacro ${M} "markdown" "Markdown document"
  !insertmacro ${M} "txt" "Text document"
  !insertmacro ${M} "log" "Log file"
  !insertmacro ${M} "pdf" "PDF document"
  !insertmacro ${M} "xlsx" "XLSX Spreadsheet"
  !insertmacro ${M} "xlsm" "XLSM Spreadsheet"
  !insertmacro ${M} "xlsb" "XLSB Spreadsheet"
  !insertmacro ${M} "xls" "XLS Spreadsheet"
  !insertmacro ${M} "xltx" "XLTX Template"
  !insertmacro ${M} "xltm" "XLTM Template"
  !insertmacro ${M} "ods" "OpenDocument spreadsheet"
  !insertmacro ${M} "fods" "OpenDocument flat spreadsheet"
  !insertmacro ${M} "numbers" "Numbers spreadsheet"
  !insertmacro ${M} "csv" "CSV table"
  !insertmacro ${M} "tsv" "Tab-separated table"
  !insertmacro ${M} "pptx" "PPTX Presentation"
  !insertmacro ${M} "pptm" "PPTM Presentation"
  !insertmacro ${M} "ppsx" "PPSX Slide Show"
  !insertmacro ${M} "ppt" "PPT Presentation"
  !insertmacro ${M} "pps" "PPS Slide Show"
  !insertmacro ${M} "odp" "OpenDocument presentation"
  !insertmacro ${M} "epub" "EPUB book"
  !insertmacro ${M} "html" "Web page"
  !insertmacro ${M} "htm" "Web page"
  !insertmacro ${M} "json" "JSON file"
  !insertmacro ${M} "xml" "XML file"
!macroend

!macro DSUnassocDesc EXT DESC
  !insertmacro DSUnassoc "${EXT}"
!macroend

!macro customInstall
  ; The app entry stays so earlier choices of "Doc Studio" keep opening, but it is hidden from Open with (the
  ; per-type entries above are offered instead) and no longer lists types of its own.
  DeleteRegKey HKCU "Software\Classes\Applications\Doc Studio.exe\SupportedTypes"
  WriteRegStr HKCU "Software\Classes\Applications\Doc Studio.exe" "NoOpenWith" ""
  WriteRegStr HKCU "Software\Classes\Applications\Doc Studio.exe" "FriendlyAppName" "Doc Studio"
  WriteRegStr HKCU "Software\Classes\Applications\Doc Studio.exe\DefaultIcon" "" '"${DS_EXE}",0'
  WriteRegStr HKCU "Software\Classes\Applications\Doc Studio.exe\shell\open\command" "" '"${DS_EXE}" "%1"'
  WriteRegStr HKCU "${DS_CAPS}" "ApplicationName" "Doc Studio"
  WriteRegStr HKCU "${DS_CAPS}" "ApplicationDescription" "View, edit and convert Markdown, PDF, documents, spreadsheets, presentations and more."
  WriteRegStr HKCU "${DS_CAPS}" "ApplicationIcon" '"${DS_EXE}",0'
  !insertmacro DSTypes DSAssoc
  WriteRegStr HKCU "Software\RegisteredApplications" "Doc Studio" "${DS_CAPS}"
  ; Settings > Installed apps: the website and the support page.
  WriteRegStr SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "URLInfoAbout" "https://griffith.software"
  WriteRegStr SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "HelpLink" "https://griffith.software/feedback.html"
  ; The Claude Code plugin starts this Doc Studio.exe, wherever it was installed.
  ${WordReplace} "$INSTDIR\Doc Studio.exe" "\" "/" "+" $R9
  FileOpen $R8 "$INSTDIR\resources\ai\claude-code\.mcp.json" w
  FileWrite $R8 '{$\r$\n  "mcpServers": {$\r$\n    "doc-studio": {$\r$\n      "type": "stdio",$\r$\n      "command": "$R9",$\r$\n      "args": ["$${CLAUDE_PLUGIN_ROOT}/server/mcp-server.js"],$\r$\n      "env": { "ELECTRON_RUN_AS_NODE": "1" }$\r$\n    }$\r$\n  }$\r$\n}$\r$\n'
  FileClose $R8
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  !insertmacro DSTypes DSUnassocDesc
  DeleteRegValue HKCU "Software\RegisteredApplications" "Doc Studio"
  DeleteRegKey HKCU "Software\Griffith Software\Doc Studio"
  DeleteRegKey /ifempty HKCU "Software\Griffith Software"
  DeleteRegKey HKCU "Software\Classes\Applications\Doc Studio.exe"
  ; The File Explorer icon handler (shell-icons.js registers it when Doc Studio starts). Its folder may still be in
  ; use by File Explorer; whatever can't be removed now goes when Explorer restarts and Doc Studio next tidies up.
  DeleteRegKey HKCU "Software\Classes\CLSID\{6C1E2B7A-4F3D-4D8E-9B21-5A7C0D3E9F41}"
  DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" "{6C1E2B7A-4F3D-4D8E-9B21-5A7C0D3E9F41}"
  RMDir /r "$LOCALAPPDATA\Doc Studio Icons"
  ; Doc Studio's temporary folders (normally removed when it quits).
  RMDir /r "$TEMP\docstudio-defaults"
  RMDir /r "$TEMP\docstudio-scan"
  RMDir /r "$TEMP\docstudio-phone"
  ${ifNot} ${isUpdated}
    ; The desktop shortcut the finish page may have made.
    Delete "$DESKTOP\${SHORTCUT_NAME}.lnk"
    ; The user's data: kept unless they choose to delete it (a silent uninstall keeps it).
    IfFileExists "$APPDATA\Doc Studio\*.*" 0 ds_data_done
      MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "Also delete your Doc Studio data?$\r$\n$\r$\nThis is your settings, recent files list, version history, backups and unsaved-work copies in $APPDATA\Doc Studio. Your documents themselves are not touched.$\r$\n$\r$\nChoose No to keep it (for example to reinstall later)." /SD IDNO IDNO ds_data_done
      RMDir /r "$APPDATA\Doc Studio"
    ds_data_done:
  ${endIf}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
