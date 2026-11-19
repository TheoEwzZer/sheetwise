// Types shared between the extension (Node) and the webview (browser).
// This file must not import any module: it is included in both bundles.

export interface BorderSide {
  /** CSS width + style, e.g. "1px solid" */
  s: string;
  /** CSS color */
  c: string;
}

/** Resolved cell style (theme colors already converted to CSS). */
export interface CellStyle {
  ff?: string; // font name
  fs?: number; // size in points
  b?: 1;
  i?: 1;
  u?: 'single' | 'double';
  st?: 1; // strikethrough
  fc?: string; // text color
  bg?: string; // CSS background (color or gradient)
  bd?: { t?: BorderSide; r?: BorderSide; b?: BorderSide; l?: BorderSide };
  ha?: string; // Excel horizontal alignment
  va?: string; // Excel vertical alignment
  wr?: 1; // wrap text
  ind?: number; // indent
  rot?: number | 'v'; // text rotation
  nf?: string; // number format (for the toolbar)
}

export interface RichRun {
  t: string;
  s?: CellStyle;
}

/** s = string, n = number, b = boolean, e = error, d = date, z = empty */
export type CellKind = 's' | 'n' | 'b' | 'e' | 'd' | 'z';

export interface CellModel {
  r: number;
  c: number;
  k: CellKind;
  t?: string; // displayed text
  e?: string; // edit text (formula with leading "=")
  v?: number; // numeric value (selection statistics)
  f?: 1; // contains a formula
  fp?: 1; // formula without a computed result
  s?: CellStyle;
  rt?: RichRun[];
  link?: string;
  note?: string;
  dv?: 1; // validation drop-down list
  m?: 1; // cell covered by a merge
}

export interface SheetMeta {
  name: string;
  tab?: string;
  hidden?: boolean;
}

export interface WorkbookMeta {
  fileName: string;
  sheets: SheetMeta[];
  warnings: string[];
  defaultFont: string;
  readonly: boolean;
}

export interface SheetModel {
  index: number;
  name: string;
  rowCount: number;
  colCount: number;
  defaultColW: number; // px
  defaultRowH: number; // px
  cols: Record<number, { w?: number; hd?: 1; s?: CellStyle }>;
  rows: Record<number, { h?: number; hd?: 1; s?: CellStyle }>;
  merges: [number, number, number, number][]; // r1, c1, r2, c2
  cells: CellModel[];
  frozen: { rows: number; cols: number };
  grid: boolean;
  zoom: number;
  hasTables: boolean;
}

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type BorderPreset = 'all' | 'outer' | 'thickOuter' | 'none' | 'top' | 'bottom' | 'left' | 'right';

export interface FormatPatch {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontName?: string;
  fontSize?: number;
  fontColor?: string | null;
  fill?: string | null;
  hAlign?: string | null;
  vAlign?: string | null;
  wrap?: boolean;
  numFmt?: string;
  border?: BorderPreset;
}

export interface Settings {
  locale: 'fr' | 'en';
  gridTheme: 'excel' | 'vscode';
}

export type StructureOp = 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';

// ---------- Messages webview -> extension ----------
export type ToHost =
  | { type: 'ready' }
  | { type: 'getSheet'; index: number }
  | { type: 'setCells'; sheet: number; items: { r: number; c: number; input: string }[] }
  | { type: 'clear'; sheet: number; ranges: CellRange[]; what: 'contents' | 'formats' | 'all' }
  | { type: 'format'; sheet: number; ranges: CellRange[]; patch: FormatPatch }
  | {
      type: 'paste';
      sheet: number;
      target: CellRange;
      rows?: string[][];
      internal?: { sheet: number; range: CellRange; cut: boolean };
    }
  | { type: 'structure'; sheet: number; op: StructureOp; at: number; count: number }
  | { type: 'merge'; sheet: number; range: CellRange; unmerge: boolean }
  | { type: 'colWidth'; sheet: number; cols: number[]; px: number }
  | { type: 'rowHeight'; sheet: number; rows: number[]; px: number }
  | { type: 'addSheet' }
  | { type: 'renameSheet'; index: number; name: string }
  | { type: 'deleteSheet'; index: number }
  | { type: 'getList'; sheet: number; r: number; c: number }
  | { type: 'openLink'; url: string }
  | { type: 'error'; message: string };

// ---------- Messages extension -> webview ----------
export type ToWebview =
  | { type: 'init'; meta: WorkbookMeta; settings: Settings }
  | { type: 'sheet'; index: number; data: SheetModel }
  | { type: 'cells'; sheet: number; cells: CellModel[]; empty: [number, number][] }
  | { type: 'reset'; meta: WorkbookMeta; activeSheet?: number }
  | { type: 'list'; r: number; c: number; values: string[] }
  | { type: 'toast'; message: string };
