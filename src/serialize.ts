// Converts the ExcelJS model into the lightweight model sent to the webview.
import type ExcelJS from 'exceljs';
import type { CellModel, CellStyle, SheetModel, BorderSide, RichRun } from './protocol.ts';
import { type ThemeInfo, resolveColor, isAutoTextColor, type ExcelColor } from './colors.ts';
import {
  type FmtCtx,
  formatNumber,
  numberEditText,
  dateToSerial,
  isDateFormat,
  normalizeFormat,
  shortDateFormat,
  booleanText,
  dateEditText,
  formatColor,
} from './numfmt.ts';

type Cell = ExcelJS.Cell;
type Worksheet = ExcelJS.Worksheet;

export const VT = {
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
} as const;

const BORDER_CSS: Record<string, string> = {
  thin: '1px solid',
  hair: '1px dotted',
  dotted: '1px dotted',
  dashed: '1px dashed',
  dashDot: '1px dashed',
  dashDotDot: '1px dashed',
  medium: '2px solid',
  mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed',
  mediumDashDotDot: '2px dashed',
  slantDashDot: '2px dashed',
  thick: '3px solid',
  double: '3px double',
};

export function colWidthToPx(w: number): number {
  return Math.max(0, Math.round(w * 7 + 5));
}
export function pxToColWidth(px: number): number {
  return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
}
export function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3);
}
export function pxToPt(px: number): number {
  return Math.round(px * 0.75 * 4) / 4;
}

export function internalRows(ws: Worksheet): (ExcelJS.Row | undefined)[] {
  return (ws as unknown as { _rows: (ExcelJS.Row | undefined)[] })._rows;
}
export function internalCells(row: ExcelJS.Row): (Cell | undefined)[] {
  return (row as unknown as { _cells: (Cell | undefined)[] })._cells;
}

export interface DataValidationsLike {
  model: Record<string, any>;
  find(address: string): unknown;
}
export function dataValidations(ws: Worksheet): DataValidationsLike {
  return (ws as unknown as { dataValidations: DataValidationsLike }).dataValidations;
}

export function richTextToString(v: unknown): string {
  if (typeof v === 'string') return v;
  if (v && typeof v === 'object' && Array.isArray((v as { richText?: unknown }).richText)) {
    return (v as { richText: { text: string }[] }).richText.map((r) => r.text).join('');
  }
  return v == null ? '' : String(v);
}

export class Serializer {
  private styleCache = new Map<string, CellStyle | undefined>();
  private theme: ThemeInfo;
  private ctx: FmtCtx;

  constructor(theme: ThemeInfo, ctx: FmtCtx) {
    this.theme = theme;
    this.ctx = ctx;
  }

  style(style: Partial<ExcelJS.Style> | undefined): CellStyle | undefined {
    if (!style) return undefined;
    let json: string;
    try {
      json = JSON.stringify(style);
    } catch {
      return undefined;
    }
    if (this.styleCache.has(json)) return this.styleCache.get(json);
    const out = this.buildStyle(style);
    this.styleCache.set(json, out);
    return out;
  }

  private fontStyle(font: Partial<ExcelJS.Font> | undefined, out: CellStyle): void {
    if (!font) return;
    if (font.name) out.ff = font.name;
    if (font.size) out.fs = font.size;
    if (font.bold) out.b = 1;
    if (font.italic) out.i = 1;
    if (font.underline) out.u = font.underline === 'double' || font.underline === 'doubleAccounting' ? 'double' : 'single';
    if (font.strike) out.st = 1;
    if (font.color && !isAutoTextColor(font.color as ExcelColor)) {
      const c = resolveColor(font.color as ExcelColor, this.theme);
      if (c) out.fc = c;
    }
  }

  private buildStyle(style: Partial<ExcelJS.Style>): CellStyle | undefined {
    const out: CellStyle = {};
    this.fontStyle(style.font, out);

    const fill = style.fill as ExcelJS.Fill | undefined;
    if (fill) {
      if (fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
        const c =
          resolveColor(fill.fgColor as ExcelColor, this.theme) ??
          resolveColor(fill.bgColor as ExcelColor, this.theme) ??
          (fill.pattern === 'solid' ? '#000000' : undefined);
        if (c) out.bg = c;
      } else if (fill.type === 'gradient') {
        const stops = (fill.stops ?? [])
          .map((s) => {
            const c = resolveColor(s.color as ExcelColor, this.theme) ?? '#FFFFFF';
            return `${c} ${Math.round(s.position * 100)}%`;
          })
          .join(', ');
        if (stops) {
          out.bg =
            fill.gradient === 'path'
              ? `radial-gradient(circle, ${stops})`
              : `linear-gradient(${((fill as ExcelJS.FillGradientAngle).degree ?? 0) + 90}deg, ${stops})`;
        }
      }
    }

    const border = style.border;
    if (border) {
      const side = (b: Partial<ExcelJS.Border> | undefined): BorderSide | undefined => {
        if (!b || !b.style) return undefined;
        const css = BORDER_CSS[b.style];
        if (!css) return undefined;
        return { s: css, c: resolveColor(b.color as ExcelColor, this.theme) ?? '#000000' };
      };
      const bd: NonNullable<CellStyle['bd']> = {};
      const t = side(border.top);
      const r = side(border.right);
      const b = side(border.bottom);
      const l = side(border.left);
      if (t) bd.t = t;
      if (r) bd.r = r;
      if (b) bd.b = b;
      if (l) bd.l = l;
      if (t || r || b || l) out.bd = bd;
    }

    const al = style.alignment;
    if (al) {
      if (al.horizontal && (al.horizontal as string) !== 'general') out.ha = al.horizontal;
      if (al.vertical) out.va = al.vertical;
      if (al.wrapText) out.wr = 1;
      if (al.indent) out.ind = al.indent;
      const rot = al.textRotation as number | 'vertical' | undefined;
      if (rot === 'vertical' || rot === 255) out.rot = 'v';
      else if (typeof rot === 'number' && rot) out.rot = rot;
    }
    if (style.numFmt && style.numFmt !== 'General') out.nf = style.numFmt;
    return Object.keys(out).length ? out : undefined;
  }

  private runs(rich: ExcelJS.RichText[]): RichRun[] {
    return rich.map((r) => {
      const s: CellStyle = {};
      this.fontStyle(r.font, s);
      return Object.keys(s).length ? { t: r.text, s } : { t: r.text };
    });
  }

  /** Fills t/e/k/v for a "simple" value (including a formula result). */
  private fillValue(out: CellModel, value: unknown, numFmt: string | undefined): void {
    const ctx = this.ctx;
    if (value === null || value === undefined) {
      out.k = 'z';
      return;
    }
    if (typeof value === 'number') {
      out.k = isDateFormat(normalizeFormat(numFmt, ctx.locale)) ? 'd' : 'n';
      out.v = value;
      out.t = formatNumber(value, numFmt, ctx);
      out.e = numberEditText(value, numFmt, ctx);
      const fc = numFmt ? formatColor(numFmt, value) : undefined;
      if (fc) out.s = { ...(out.s ?? {}), fc };
      return;
    }
    if (typeof value === 'string') {
      out.k = 's';
      out.t = value;
      out.e = value;
      return;
    }
    if (typeof value === 'boolean') {
      out.k = 'b';
      out.t = booleanText(value, ctx.locale);
      out.e = out.t;
      return;
    }
    if (value instanceof Date) {
      const serial = dateToSerial(value, ctx.date1904);
      const fmt = isDateFormat(normalizeFormat(numFmt, ctx.locale)) ? numFmt : shortDateFormat(ctx.locale);
      out.k = 'd';
      out.v = serial;
      out.t = Number.isNaN(serial) ? '#VALUE!' : formatNumber(serial, fmt, ctx);
      out.e = Number.isNaN(serial) ? '' : dateEditText(serial, ctx);
      return;
    }
    if (typeof value === 'object') {
      const o = value as Record<string, unknown>;
      if (typeof o.error === 'string') {
        out.k = 'e';
        out.t = o.error;
        out.e = o.error;
        return;
      }
      if (Array.isArray(o.richText)) {
        out.k = 's';
        out.rt = this.runs(o.richText as ExcelJS.RichText[]);
        out.t = richTextToString(o);
        out.e = out.t;
        return;
      }
      if (typeof o.hyperlink === 'string') {
        out.k = 's';
        out.link = o.hyperlink;
        const text = o.text as unknown;
        if (text && typeof text === 'object' && Array.isArray((text as { richText?: unknown }).richText)) {
          out.rt = this.runs((text as { richText: ExcelJS.RichText[] }).richText);
        }
        out.t = richTextToString(text ?? o.hyperlink);
        out.e = out.t;
        return;
      }
    }
    out.k = 's';
    out.t = String(value);
    out.e = out.t;
  }

  cell(ws: Worksheet, cell: Cell, withValidation: boolean): CellModel | undefined {
    const r = Number(cell.row);
    const c = Number(cell.col);
    const out: CellModel = { r, c, k: 'z' };
    const s = this.style(cell.style);
    if (s) out.s = s;
    if (cell.type === VT.Merge) {
      out.m = 1;
      return out;
    }
    const numFmt = cell.numFmt;
    const value = cell.value as unknown;
    if (cell.type === VT.Formula) {
      out.f = 1;
      const result = cell.result as unknown; // cell.value hides 0 / "" / FALSE
      if (result === undefined) {
        out.fp = 1;
        out.k = 's';
      } else {
        this.fillValue(out, result, numFmt);
      }
      out.e = '=' + (cell.formula ?? '');
    } else {
      this.fillValue(out, value, numFmt);
    }
    const note = cell.note as unknown;
    if (note) {
      out.note =
        typeof note === 'string'
          ? note
          : ((note as { texts?: { text: string }[] }).texts ?? []).map((t) => t.text).join('');
    }
    if (withValidation) {
      const dv = dataValidations(ws).find(cell.address) as { type?: string } | undefined;
      if (dv && dv.type === 'list') out.dv = 1;
    }
    if (out.k === 'z' && !out.s && !out.note && !out.dv && !out.f) return undefined;
    return out;
  }

  sheet(ws: Worksheet, index: number): SheetModel {
    const props = ws.properties ?? ({} as ExcelJS.WorksheetProperties);
    const defaultColW = colWidthToPx(props.defaultColWidth ?? 8.43);
    const defaultRowH = ptToPx(props.defaultRowHeight ?? 15);
    const cols: SheetModel['cols'] = {};
    const columns = (ws as unknown as { _columns?: ExcelJS.Column[] })._columns ?? [];
    columns.forEach((col, i) => {
      if (!col) return;
      const entry: SheetModel['cols'][number] = {};
      if (col.width !== undefined && col.width !== null) entry.w = colWidthToPx(col.width);
      if (col.hidden) entry.hd = 1;
      const st = this.style(col.style as Partial<ExcelJS.Style>);
      if (st && (st.bg || st.bd || st.ff || st.fs || st.b || st.ha)) entry.s = st;
      if (Object.keys(entry).length) cols[i + 1] = entry;
    });

    const rows: SheetModel['rows'] = {};
    const cells: CellModel[] = [];
    const hasValidations = Object.keys(dataValidations(ws).model ?? {}).length > 0;
    let maxRow = 0;
    let maxCol = columns.length;
    internalRows(ws).forEach((row) => {
      if (!row) return;
      const rn = row.number;
      const entry: SheetModel['rows'][number] = {};
      if (row.height) entry.h = ptToPx(row.height);
      if (row.hidden) entry.hd = 1;
      const st = this.style((row as unknown as { style?: Partial<ExcelJS.Style> }).style);
      if (st && (st.bg || st.bd || st.ff || st.fs || st.b)) entry.s = st;
      if (Object.keys(entry).length) rows[rn] = entry;
      internalCells(row).forEach((cell) => {
        if (!cell) return;
        const m = this.cell(ws, cell, hasValidations);
        if (m) {
          cells.push(m);
          if (m.k !== 'z' || m.s?.bg || m.s?.bd) {
            maxRow = Math.max(maxRow, rn);
            maxCol = Math.max(maxCol, m.c);
          }
        }
      });
    });

    const merges: SheetModel['merges'] = [];
    const mergeMap = (ws as unknown as { _merges: Record<string, { model: { top: number; left: number; bottom: number; right: number } }> })._merges;
    for (const key of Object.keys(mergeMap)) {
      const m = mergeMap[key].model;
      merges.push([m.top, m.left, m.bottom, m.right]);
      maxRow = Math.max(maxRow, m.bottom);
      maxCol = Math.max(maxCol, m.right);
    }

    const view = (ws.views?.[0] ?? {}) as Partial<ExcelJS.WorksheetViewFrozen> & Partial<ExcelJS.WorksheetViewCommon>;
    const frozen =
      view.state === 'frozen' ? { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 } : { rows: 0, cols: 0 };

    return {
      index,
      name: ws.name,
      rowCount: Math.max(maxRow, ws.rowCount),
      colCount: maxCol,
      defaultColW,
      defaultRowH,
      cols,
      rows,
      merges,
      cells,
      frozen,
      grid: view.showGridLines !== false,
      zoom: (view as { zoomScale?: number }).zoomScale ?? 100,
      hasTables: Object.keys((ws as unknown as { tables?: object }).tables ?? {}).length > 0,
    };
  }
}
