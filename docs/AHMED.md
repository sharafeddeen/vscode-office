# Ahmed Spreadsheet Viewer

A local VS Code fork of cweijan/vscode-office, maintained at https://github.com/sharafeddeen/vscode-office on branch `ahmed/main`. The upstream source and MIT attribution remain in this repository. Only the spreadsheet provider is activated by this fork.

## Install and use

Requirements: VS Code 1.85+, Node.js 24+ for development. LibreOffice and a Python interpreter with UNO bindings are required for native workbook editing. On Windows the normal LibreOffice installation includes the runtime; Linux generally needs libreoffice-calc and python3-uno. Saved-result viewing does not require a backend.

```bash
npm ci --no-audit --no-fund
npm test
npm run build
npm run package
code --install-extension ahmed-spreadsheet-viewer-0.2.0.vsix --force
```

Reload VS Code. For a spreadsheet, select **Reopen Editor With → Ahmed Spreadsheet Viewer**. The extension ID is `sharafeddeen.ahmed-spreadsheet-viewer`; it is independent of the Marketplace extension. Existing Markdown, PDF, themes and other editor associations are untouched.

Version 0.2.0 keeps the upstream spreadsheet toolbar and connects workbook editing to a persistent LibreOffice document through Python UNO. Each open workbook has a private headless process, a named-pipe connection and a short temporary directory. The original file is copied into that directory; ordinary edits modify the native document. Save, Save As, backup and revert use that document rather than rebuilding an XLSX from visible cells.

Connected edits: cell values/formulas, native numeric/date input parsing, text-format values, fonts, fills, text colors, borders, alignment/wrapping, number formats, ordinary comments, merges, column widths and row heights. Row/column insert/delete and sheet add/delete/rename/duplicate/reorder run in LibreOffice, which updates its references and named ranges. Toolbar undo/redo use LibreOffice’s document history. All imported sheets have tabs, including hidden sheets; the native document retains their visibility.

Changing content invalidates displayed formula caches. Edits are synchronized after a short pause and automatically recalculated by default. Explicit **Recalculate** calls calculateAll on the same document. Complete native output ranges are imported, including array-result cells. Derived array cells are read only in the grid; array creation/resizing/editing is not a complete Excel-equivalent interface.

Writable XLSX/XLS/ODS workbooks open in native editing mode in trusted workspaces when the runtime is available. An unavailable/disabled backend or an untrusted workspace leaves workbooks read only with an explanation. XLSM remains read only to avoid promising VBA preservation. CSV/TSV retain the upstream text editor/save path.

VS Code tracks unsaved changes and prompts on close. Source conflicts are checked before overwrite. Native edits never write the source until Save (including VS Code Auto Save if enabled). Save As supports XLSX, XLS and ODS through LibreOffice’s export filters. File-format fidelity is LibreOffice’s import/export fidelity; unsupported Excel-specific features are not guaranteed.

Native metadata preservation is tested for named ranges and conditional formatting. It is not a claim of byte-identical round trips or exhaustive preservation of every OOXML extension.

The following interface operations are not connected in this phase: adding/removing images, editing validation/protection/hyperlinks/freeze/filter definitions, and dragging whole rows/columns to reorder them. Unchanged existing features stay in the native document. An unsupported edit is rejected explicitly; it does not silently switch to the old grid exporter. **Reload native state** restores the latest accepted native document to the interface. Use the desktop application for these operations and for charts/pivot editing.

## Values and calculation

Saved results are displayed without running the canvas widget's limited JavaScript formula evaluator. The formula bar retains the expression. Zero, false, an empty string, a real saved spreadsheet error, and a missing result are distinct. Missing results display **Not calculated**, never a made-up zero or `#NAME?`.

Saved values can be stale; the label says **Saved results**, not verified/current. **Recalculate** explicitly refreshes through the native spreadsheet document. With the default settings, missing results trigger calculation automatically in a trusted workspace. A disabled/unavailable backend leaves an honest saved-result preview.

Settings:
- `ahmedSpreadsheet.backend`: `auto` (LibreOffice UNO editing when available), `excel` (read-only calculation fallback), `libreoffice`, or `none`.
- `ahmedSpreadsheet.libreOfficePath`: optional absolute path to the installed executable.
- `ahmedSpreadsheet.pythonPath`: optional path to a Python interpreter with UNO.
- `ahmedSpreadsheet.autoRecalculate`: defaults to true; recalculates after edits.
- `ahmedSpreadsheet.autoCalculateMissing`: defaults to true.
- `ahmedSpreadsheet.timeoutSeconds`: defaults to 90.

Backend executable settings are machine-scoped. Native work requires workspace trust. The worker opens documents with NEVER_EXECUTE macro mode and NO_UPDATE external-link mode. It communicates with its own isolated LibreOffice process over a random named pipe. Cancellation/timeouts terminate that worker’s process tree. The latest accepted native state is checkpointed as ODS through an atomic replacement, allowing reconnection after a stopped worker. Interrupted structural operations may require **Reload native state** before editing resumes.

Closing the document disposes its worker and removes its own temporary files. Backups use native document exports. Runtime/environment settings from Electron are excluded from subprocesses. No workbook data is sent to a service.

Excel automation and the earlier conversion-based calculator remain available for read-only calculation fallback. Their results cache is separate; **Ahmed Spreadsheet Viewer: Clear calculation cache** removes that legacy cache. Excel-specific fallback remains untested on machines without Excel.

## Import and rendering

The package normalizer follows OOXML relationship types and resolves absolute/relative paths; it canonicalizes comments/VML on a disposable ZIP for ExcelJS. It does not rewrite source files. Missing referenced parts raise useful errors. SheetJS is a recovery reader for variants the primary reader cannot interpret; fallback use is disclosed.

Supported preview features include sheet names/order, hidden-sheet access, cells, saved results, formulas, ordinary legacy comments, merges, number formats, fonts, fills, borders, alignment/wrapping, row heights, column widths, and the upstream grid/image support.

Compatibility boundaries:
| Feature | Behavior |
| --- | --- |
| XLSX/XLSM worksheet data | Primary supported preview path; macros never run |
| XLS/ODS | Upstream initial preview; native editing/export and XLSX projections through LibreOffice |
| CSV/TSV | Upstream text-oriented editing/save path |
| Chart objects | Not rendered; disclosed in workbook details |
| Pivot tables | Saved worksheet values shown; no pivot editing/refresh |
| Threaded comments | Disclosed, not rendered |
| External links | Saved results only; links are not refreshed |
| Conditional formatting | Not an Excel-equivalent rendering engine; use desktop app for exact visual validation |
| Editing/round-trip saving | Native LibreOffice document; retains supported document features within its format filters |
| Password-encrypted/corrupt files | Explicit loading error and desktop-app action |

## Tests

```bash
npm test
RUN_UNO=1 RUN_NATIVE=1 npm test
# Optionally select the installed native backend:
RUN_NATIVE=1 NATIVE_BACKEND=libreoffice npm test
npm run build
npm run test:integration
npm run test:corpus -- /absolute/path/to/external/workbook/directory
```

Unit tests generate synthetic workbooks with arbitrary comment paths, relative/absolute targets, hidden sheets, formatting, cross-sheet formulas, cached errors, false/zero/empty values, and missing results. They check untouched source bytes and worksheet XML, plus cancellation, invalid packages, edit invalidation and save/reopen fidelity for supported values and formatting.

RUN_UNO tests require LibreOffice/Python UNO and fail if requested but unavailable. They cover native named ranges, formulas, formatting, arrays, undo/redo, insertion, export/reopen and temporary-file cleanup. Integration tests launch an isolated VS Code profile and verify the actual webview/canvas, saved results, missing-result display, recalculated results, the full toolbar, formula-bar edits, undo/redo, formatting, find/replace, saving, Save As and revert. Set `VSCODE_EXECUTABLE` to use a particular Code executable. Set `AHMED_ORIGINAL_WORKBOOK` to additionally exercise an external workbook with six sheets and `QUICK_SPLITS!H5 = OK`; that workbook is never copied into Git.

Corpus verification reads external XLSX files, checks sheet identity, result rendering and source hashes, and writes a local ignored `.cache/corpus-report.json`. No customer workbooks or corpus reports belong in the public fork or release package.

CI builds/tests/packages synthetic fixtures on Windows and Linux, with a separate Linux job provisioning LibreOffice/Python UNO for native tests. Real VS Code webview/native checks run locally on Windows.

## Repository relationship and updates

The private parent `sharafeddeen/vscode-extensions` includes this fork as `forks/vscode-office` and pins a tested commit. This repository's `origin` is your GitHub fork; `upstream` is the original repository. Commit/push child changes first, then commit/push the parent's new submodule pointer. Work on a named branch; a recursive clone initially checks out the pinned commit with detached HEAD.

Fetch upstream and merge selected changes into `ahmed/main`, run tests/build/corpus checks, package/install the new version, then advance the parent pointer. Retain the lockfile for repeatable dependency installation. Release VSIX artifacts belong on GitHub Releases or local ignored storage, not in Git source history. Roll back by installing an earlier VSIX and pinning its fork commit in the parent.
