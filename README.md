# Sheetwise — XLSX Viewer & Editor for VS Code

[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![VS Code](https://img.shields.io/badge/VS%20Code-%5E1.85-007ACC.svg?logo=visualstudiocode)](https://code.visualstudio.com/)

Open `.xlsx` workbooks directly in VS Code **with their real formatting**, edit them like you would in Excel, and save with <kbd>Ctrl</kbd>+<kbd>S</kbd>.

No more opening Excel just to check a value, fix a typo or tweak a formula: Sheetwise renders fonts, colors, borders, merged cells, number formats and frozen panes, recalculates formulas as you type, and plugs into VS Code's native undo/redo, dirty state and hot exit.

## Features

### Faithful rendering

- **Fonts & text** — font family and size, bold / italic / underline / strikethrough, colors (including theme colors and tints), rich text
- **Cells** — solid and gradient fills, borders (thin, medium, thick, dashed, dotted, double), horizontal / vertical alignment, indent, wrap text, vertical text
- **Layout** — merged cells, column widths, row heights, hidden rows / columns, frozen panes, gridlines, sheet zoom, tab colors, hidden sheets (greyed out)
- **Number formats** — every Excel format: currency, percentage, dates, times, scientific, conditional colors (`[Red]`…), in French or English conventions
- **Extras** — hyperlinks (<kbd>Ctrl</kbd>+click), notes (red corner + tooltip), data validation dropdown lists, text overflowing into empty neighbouring cells just like Excel
- **Fast** — the grid is rendered progressively, so large sheets open quickly

### Editing

| Action | How |
| --- | --- |
| Enter data | Just type, or <kbd>F2</kbd> / double-click to edit, or use the formula bar |
| Confirm / cancel | <kbd>Enter</kbd>, <kbd>Tab</kbd> (with <kbd>Shift</kbd> to go backwards) / <kbd>Esc</kbd>. <kbd>Alt</kbd>+<kbd>Enter</kbd> inserts a line break |
| Formulas | `=SUM(A1:A3)` — French function names and `;` separators also work (`=SOMME(A1;A3)`) |
| Navigate | Arrows, <kbd>Ctrl</kbd>+arrows, <kbd>Home</kbd>, <kbd>Ctrl</kbd>+<kbd>Home</kbd>, <kbd>Ctrl</kbd>+<kbd>End</kbd>, <kbd>PgUp</kbd> / <kbd>PgDn</kbd> |
| Select | <kbd>Shift</kbd>+arrows, drag, click headers, <kbd>Ctrl</kbd>+<kbd>Space</kbd> / <kbd>Shift</kbd>+<kbd>Space</kbd>, <kbd>Ctrl</kbd>+<kbd>A</kbd> |
| Copy / cut / paste | <kbd>Ctrl</kbd>+<kbd>C</kbd> / <kbd>X</kbd> / <kbd>V</kbd> — Excel-compatible TSV; formulas and formats are kept within the workbook |
| Fill down / right | <kbd>Ctrl</kbd>+<kbd>D</kbd> / <kbd>Ctrl</kbd>+<kbd>R</kbd> |
| Clear | <kbd>Delete</kbd> (contents), context menu (formats, all) |
| Format | Toolbar: font, size, bold / italic / underline / strikethrough (<kbd>Ctrl</kbd>+<kbd>B</kbd> / <kbd>I</kbd> / <kbd>U</kbd> / <kbd>5</kbd>), colors, alignment, wrap text, merge, borders, number format |
| Rows / columns | Right-click to insert / delete — formulas, merges, data validations and conditional formats are shifted accordingly |
| Width / height | Drag a header edge; double-click to autofit the column |
| Sheets | Tabs: click to switch, double-click to rename, right-click to insert / delete, `+` button to add |
| Undo / redo | <kbd>Ctrl</kbd>+<kbd>Z</kbd> / <kbd>Ctrl</kbd>+<kbd>Y</kbd> (VS Code's own history) |
| Find | <kbd>Ctrl</kbd>+<kbd>F</kbd> |
| Zoom | Zoom dropdown or <kbd>Ctrl</kbd>+mouse wheel |

### Formula recalculation

Formulas that depend on edited cells are recalculated incrementally, using [`fast-formula-parser`](https://github.com/LesterLyu/fast-formula-parser) extended with `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `TEXTJOIN`, `XLOOKUP`, `SWITCH` and `VALUE`.

Untouched formulas keep the value Excel last computed. If a function isn't supported, the last known value is kept. Saved files are flagged so that Excel fully recalculates them on open.

## Installation

Sheetwise is not on the Marketplace yet. Build the VSIX from source and install it:

```bash
git clone https://github.com/TheoEwzZer/sheetwise.git
cd sheetwise
npm install
npm run package
code --install-extension dist/sheetwise.vsix
```

Then open any `.xlsx` file — Sheetwise is registered as its default editor. To switch back to another editor, use **Reopen Editor With…** from the editor title menu.

## Settings

| Setting | Values | Description |
| --- | --- | --- |
| `xlsxEditor.locale` | `auto` (default), `fr`, `en` | Conventions used to display and enter numbers, dates and booleans. `auto` follows the VS Code display language |
| `xlsxEditor.gridTheme` | `excel` (default), `vscode` | `excel`: white grid like Excel. `vscode`: cells without a fill use the VS Code theme colors |
| `xlsxEditor.recalculate` | `true` (default), `false` | Recalculate formulas that depend on edited cells |

Command: **XLSX: New Excel Workbook (.xlsx)**.

## Known limitations

Saving rewrites the file through [ExcelJS](https://github.com/exceljs/exceljs). The following are **lost on save**: charts, pivot tables, slicers, macros, shapes / text boxes and form controls. A warning banner is shown when a file contains any of them — don't save it from Sheetwise in that case (or work on a copy).

Also:

- Conditional formatting is preserved but not rendered
- Images are not displayed
- Defined names and external references are not fully recalculated
- Inserting / deleting columns inside an Excel Table object is refused

## Development

```bash
npm install
npm run build      # bundle the extension and the webview (esbuild)
npm run watch      # rebuild on change
npm run typecheck
npm test           # unit and integration tests (node:test)
npm run package    # build dist/sheetwise.vsix
```

Press <kbd>F5</kbd> in VS Code to launch an Extension Development Host with the extension loaded.

### Project structure

| Path | Role |
| --- | --- |
| `src/extension.ts` | Activation, command registration |
| `src/provider.ts` | `CustomEditorProvider` and messaging with the webview |
| `src/xlsxDocument.ts` | Document model: loading, edit operations, undo/redo, save, backups |
| `src/serialize.ts` | ExcelJS → view model (resolved styles, formatted text) |
| `src/recalc.ts`, `src/formulaFunctions.ts` | Incremental formula recalculation and extra functions |
| `src/refs.ts`, `src/structure.ts` | Reference rewriting when inserting / deleting rows and columns |
| `src/numfmt.ts`, `src/parseInput.ts`, `src/colors.ts` | Number formats, input parsing, color resolution |
| `src/protocol.ts` | Message types shared by the extension and the webview |
| `webview/` | HTML grid: progressive rendering, selection, editing, clipboard, menus |
| `test/` | `node:test` test suites |

## Contributing

Issues and pull requests are welcome. Please run `npm run typecheck` and `npm test` before submitting a PR.

## License

[MIT](LICENSE). Bundled third-party libraries (ExcelJS, SSF, fast-formula-parser…) and their licenses are listed in [`ThirdPartyNotices.txt`](ThirdPartyNotices.txt).
