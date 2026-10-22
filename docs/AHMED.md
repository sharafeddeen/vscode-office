# Ahmed Spreadsheet Viewer

A local VS Code fork of cweijan/vscode-office, maintained at https://github.com/sharafeddeen/vscode-office on branch `ahmed/main`. The upstream source and MIT attribution remain in this repository. Only the spreadsheet provider is activated by this fork.

## Install and use

Requirements: VS Code 1.85+, Node.js 24+ for development. Excel or LibreOffice is needed only to calculate formula results that are missing or to refresh results.

```bash
npm ci --no-audit --no-fund
npm test
npm run build
npm run package
code --install-extension ahmed-spreadsheet-viewer-0.1.1.vsix --force
```

Reload VS Code. For a spreadsheet, select **Reopen Editor With → Ahmed Spreadsheet Viewer**. The extension ID is `sharafeddeen.ahmed-spreadsheet-viewer`; it is independent of the Marketplace extension. Existing Markdown, PDF, themes and other editor associations are untouched.

Version 0.1.1 restores the complete upstream spreadsheet interface: formatting toolbar, editable cells and formula bar, undo/redo, clipboard, find/replace, Save/Save As, sheet operations, auto-fit, context menus and light/adaptive theme controls. There is one formula bar. Recalculation, zoom, comment inspection and desktop-app actions live in a compact status bar below the workbook. All imported sheets have tabs, including hidden sheets; their original visibility state is preserved when exported.

Ordinary writable files open in edit mode. VS Code tracks unsaved changes, prompts on close, and supports save, revert and backup. Disk changes are checked before overwriting. Read-only resources and XLSM files open in read mode with Save As available; the XLSX exporter cannot preserve VBA projects. For CSV/TSV, the upstream text-editor action remains available.

Changing cell content or structure invalidates formula caches throughout the workbook, including cross-sheet results. Formatting alone retains results. **Recalculate** uses the current unsaved grid when there are edits. Its results stay in memory until you save. Copying cells copies displayed values. Saving/exporting preserves formulas and available results, number formats, ordinary comments and hidden-sheet state in the supported grid model.

## Values and calculation

Saved results are displayed without running the canvas widget's limited JavaScript formula evaluator. The formula bar retains the expression. Zero, false, an empty string, a real saved spreadsheet error, and a missing result are distinct. Missing results display **Not calculated**, never a made-up zero or `#NAME?`.

Saved values can be stale; the label says **Saved results**, not verified/current. **Recalculate** explicitly refreshes through a native spreadsheet application. With the default settings, missing results trigger calculation automatically in a trusted workspace. A disabled/unavailable backend leaves an honest saved-result preview.

Settings:
- `ahmedSpreadsheet.backend`: `auto` (Excel on Windows, then LibreOffice), `excel`, `libreoffice`, or `none`.
- `ahmedSpreadsheet.libreOfficePath`: optional absolute path to the installed executable.
- `ahmedSpreadsheet.autoCalculateMissing`: defaults to true.
- `ahmedSpreadsheet.timeoutSeconds`: defaults to 90.

Backend executable settings are machine-scoped. Recalculation requires workspace trust. Native workers operate on extension-owned temporary copies, disable macros, and disable external-link refresh. Calculation never saves the original workbook. The editor writes a workbook only through the normal Save/Save As lifecycle (including VS Code Auto Save if you enable it). Results are cached in VS Code's extension global storage by source hash, backend identity/version (binary fingerprint on Windows) and cache format version. Explicit recalculation bypasses the cache. **Ahmed Spreadsheet Viewer: Clear calculation cache** removes only that extension-owned cache.

LibreOffice uses an isolated profile per job under a short OS-temp path, avoiding native Windows path limits. Excel uses a separate automation instance and records its owned process for cancellation cleanup. Process arguments are passed as arrays rather than constructed shell commands. A timeout/cancellation terminates the job and cleans its temporary files. Excel support requires a registered Excel automation server; it cannot be validated on machines without Excel.

For an unedited workbook, only calculated values are overlaid onto the original preview. A backend's rewritten workbook is not substituted for the source's styles, formulas, comments or layout. Recalculating edited grids exports their supported cells/styles to a temporary XLSX; advanced features such as named ranges outside the grid export model require the desktop application. A native engine may implement some Excel formulas differently; the UI identifies the engine used. Macros, add-in functions, missing external dependencies and encrypted workbooks may need the desktop application.

## Import and rendering

The package normalizer follows OOXML relationship types and resolves absolute/relative paths; it canonicalizes comments/VML on a disposable ZIP for ExcelJS. It does not rewrite source files. Missing referenced parts raise useful errors. SheetJS is a recovery reader for variants the primary reader cannot interpret; fallback use is disclosed.

Supported preview features include sheet names/order, hidden-sheet access, cells, saved results, formulas, ordinary legacy comments, merges, number formats, fonts, fills, borders, alignment/wrapping, row heights, column widths, and the upstream grid/image support.

Compatibility boundaries:
| Feature | Behavior |
| --- | --- |
| XLSX/XLSM worksheet data | Primary supported preview path; macros never run |
| XLS/ODS/CSV/TSV | Upstream recovery reader; XLSX cache inspection/native overlay is not applied |
| Chart objects | Not rendered; disclosed in workbook details |
| Pivot tables | Saved worksheet values shown; no pivot editing/refresh |
| Threaded comments | Disclosed, not rendered |
| External links | Saved results only; links are not refreshed |
| Conditional formatting | Not an Excel-equivalent rendering engine; use desktop app for exact visual validation |
| Editing/round-trip saving | Upstream grid export; advanced OOXML features outside its model may be lost. Use Save As or the desktop app for complex workbooks |
| Password-encrypted/corrupt files | Explicit loading error and desktop-app action |

## Tests

```bash
npm test
RUN_NATIVE=1 npm test
# Optionally select the installed native backend:
RUN_NATIVE=1 NATIVE_BACKEND=libreoffice npm test
npm run build
npm run test:integration
npm run test:corpus -- /absolute/path/to/external/workbook/directory
```

Unit tests generate synthetic workbooks with arbitrary comment paths, relative/absolute targets, hidden sheets, formatting, cross-sheet formulas, cached errors, false/zero/empty values, and missing results. They check untouched source bytes and worksheet XML, plus cancellation, invalid packages, edit invalidation and save/reopen fidelity for supported values and formatting.

Native tests require a backend and fail if requested but unavailable. Integration tests launch an isolated VS Code profile and verify the actual webview/canvas, saved results, missing-result display, recalculated results, the full toolbar, formula-bar edits, undo/redo, formatting, find/replace, saving, Save As and revert. Set `VSCODE_EXECUTABLE` to use a particular Code executable. Set `AHMED_ORIGINAL_WORKBOOK` to additionally exercise an external workbook with six sheets and `QUICK_SPLITS!H5 = OK`; that workbook is never copied into Git.

Corpus verification reads external XLSX files, checks sheet identity, result rendering and source hashes, and writes a local ignored `.cache/corpus-report.json`. No customer workbooks or corpus reports belong in the public fork or release package.

CI builds/tests/packages synthetic fixtures on Windows and Linux. Native/UI integration checks are local unless the corresponding tools are provisioned.

## Repository relationship and updates

The private parent `sharafeddeen/vscode-extensions` includes this fork as `forks/vscode-office` and pins a tested commit. This repository's `origin` is your GitHub fork; `upstream` is the original repository. Commit/push child changes first, then commit/push the parent's new submodule pointer. Work on a named branch; a recursive clone initially checks out the pinned commit with detached HEAD.

Fetch upstream and merge selected changes into `ahmed/main`, run tests/build/corpus checks, package/install the new version, then advance the parent pointer. Retain the lockfile for repeatable dependency installation. Release VSIX artifacts belong on GitHub Releases or local ignored storage, not in Git source history. Roll back by installing an earlier VSIX and pinning its fork commit in the parent.
