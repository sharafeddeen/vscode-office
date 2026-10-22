# Ahmed Spreadsheet Viewer

The complete Office Viewer spreadsheet interface inside VS Code, with workbook compatibility fixes and a persistent LibreOffice document for edits, calculations and saving.

This is Ahmed's independently identified fork of [cweijan/vscode-office](https://github.com/cweijan/vscode-office). Its extension ID is `sharafeddeen.ahmed-spreadsheet-viewer`. Only its spreadsheet provider is registered.

See [the setup, compatibility, calculation and testing guide](docs/AHMED.md).

```bash
npm ci
npm test
npm run build
npm run package
code --install-extension ahmed-spreadsheet-viewer-0.2.0.vsix --force
```

Reload VS Code, then use **Reopen Editor With → Ahmed Spreadsheet Viewer**. Saved-result viewing works without LibreOffice. Workbook editing requires LibreOffice and Python UNO; Windows normally uses LibreOffice’s bundled Python. CSV/TSV editing remains independent.

The original project remains available upstream; its source files, history and MIT license are retained.

The upstream toolbar, editing, undo/redo, formatting, find/replace, Save/Save As and theme controls are retained. The calculation/status controls appear below the workbook. See the guide for editing/export limitations.
