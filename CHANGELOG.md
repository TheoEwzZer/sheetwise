# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-11-19

### Added

- Formatted rendering of `.xlsx` workbooks: fonts, colors, fills, borders, alignment, merged cells, column widths / row heights, frozen panes, number formats, rich text, hyperlinks, notes and data validation lists.
- Cell editing with a formula bar, keyboard navigation and selection, Excel-compatible copy / cut / paste, fill down / right.
- Formatting toolbar: font, size, bold / italic / underline / strikethrough, text and fill colors, alignment, wrap text, merge, borders, number formats.
- Insert / delete rows and columns, with formulas, merges, data validations and conditional formats shifted accordingly.
- Sheet management: add, rename, delete, switch between sheets.
- Incremental formula recalculation, including `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `TEXTJOIN`, `XLOOKUP`, `SWITCH` and `VALUE`.
- Native VS Code undo / redo, save, and backups (hot exit).
- `XLSX: New Excel Workbook (.xlsx)` command.
- Settings: `xlsxEditor.locale`, `xlsxEditor.gridTheme`, `xlsxEditor.recalculate`.
- Warning banner when a file contains content that would be lost on save (charts, pivot tables, macros…).

[0.1.0]: https://github.com/TheoEwzZer/sheetwise/releases/tag/v0.1.0
