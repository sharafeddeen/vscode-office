# Ahmed Spreadsheet Viewer

Read-only spreadsheet previews inside VS Code, with faithful saved formula results and native recalculation on temporary copies.

This is Ahmed's independently identified fork of [cweijan/vscode-office](https://github.com/cweijan/vscode-office). Its extension ID is `sharafeddeen.ahmed-spreadsheet-viewer`. Only its spreadsheet provider is registered.

See [the setup, compatibility, calculation and testing guide](docs/AHMED.md).

```bash
npm ci
npm test
npm run build
npm run package
code --install-extension ahmed-spreadsheet-viewer-0.1.0.vsix --force
```

Reload VS Code, then use **Reopen Editor With → Ahmed Spreadsheet Viewer**. Excel/LibreOffice is optional for saved-result viewing and required for native recalculation.

The original project remains available upstream; its source files, history and MIT license are retained.
