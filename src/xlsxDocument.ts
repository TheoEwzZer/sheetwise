import * as vscode from 'vscode';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import type {
  CellModel,
  CellRange,
  FormatPatch,
  StructureOp,
  ToWebview,
  WorkbookMeta,
  SheetModel,
} from './protocol.ts';
import { DEFAULT_THEME, type ThemeInfo, parseTheme, resolveColor, cssToArgb, type ExcelColor } from './colors.ts';
import type { Locale } from './numfmt.ts';
import { parseInput } from './parseInput.ts';
import { Recalc, cellKey } from './recalc.ts';
import { Serializer, VT, pxToColWidth, pxToPt, richTextToString, dataValidations } from './serialize.ts';
import { deleteSheet, renameSheet, spliceSheet, unshareFormulas, validateSheetName } from './structure.ts';
import { offsetRef, parseRangeText, transformRefs, address, rangeToA1 } from './refs.ts';

export interface DocSettings {
  locale: Locale;
  recalc: boolean;
}

interface CellSnap {
  value: unknown;
  style: unknown;
  skipValue: boolean;
}

interface CellChange {
  s: number;
  r: number;
  c: number;
  before: CellSnap;
  after: CellSnap;
}

export interface EditEvent {
  label: string;
  undo(): Promise<void>;
  redo(): Promise<void>;
}

const MAX_CELLS_PER_OP = 500_000;

function clone<T>(v: T): T {
  return v === undefined ? v : structuredClone(v);
}

function snap(cell: ExcelJS.Cell): CellSnap {
  const skipValue = cell.type === VT.Merge;
  let value: unknown = skipValue ? undefined : cell.value;
  // cell.value hides falsy formula results (0, "", FALSE)
  if (cell.type === VT.Formula) value = { ...(value as object), result: cell.result };
  return { value: clone(value), style: clone(cell.style ?? {}), skipValue };
}

/** ExcelJS shares style objects between cells: always replace the whole object. */
function setStyle<K extends keyof ExcelJS.Style>(cell: ExcelJS.Cell, key: K, value: ExcelJS.Style[K] | undefined): void {
  const next = { ...(cell.style ?? {}) } as Record<string, unknown>;
  if (value === undefined) delete next[key];
  else next[key] = value;
  cell.style = next as Partial<ExcelJS.Style>;
}

const THIN: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } };
const MEDIUM: Partial<ExcelJS.Border> = { style: 'medium', color: { argb: 'FF000000' } };

export class XlsxDocument implements vscode.CustomDocument {
  wb!: ExcelJS.Workbook;
  private theme: ThemeInfo = DEFAULT_THEME;
  private serializer!: Serializer;
  private recalc!: Recalc;
  private warnings: string[] = [];
  private queue: Promise<unknown> = Promise.resolve();
  private lastSave = 0;
  dirty = false;

  private readonly _onDidEdit = new vscode.EventEmitter<EditEvent>();
  readonly onDidEdit = this._onDidEdit.event;
  private readonly _onDidUpdate = new vscode.EventEmitter<ToWebview>();
  readonly onDidUpdate = this._onDidUpdate.event;
  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  readonly onDidDispose = this._onDidDispose.event;
  private watcher: vscode.FileSystemWatcher | undefined;

  private constructor(
    readonly uri: vscode.Uri,
    private settings: DocSettings,
  ) {}

  static async create(uri: vscode.Uri, backupId: string | undefined, settings: DocSettings): Promise<XlsxDocument> {
    const doc = new XlsxDocument(uri, settings);
    const source = backupId ? vscode.Uri.parse(backupId) : uri;
    let bytes: Uint8Array;
    try {
      bytes = await vscode.workspace.fs.readFile(source);
    } catch (err) {
      if (backupId) bytes = await vscode.workspace.fs.readFile(uri);
      else throw err;
    }
    if (bytes.byteLength === 0) bytes = await XlsxDocument.emptyWorkbookBytes();
    await doc.load(bytes, true);
    if (backupId) doc.dirty = true;
    doc.watch();
    return doc;
  }

  static async emptyWorkbookBytes(): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1');
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }

  private watch(): void {
    if (this.uri.scheme !== 'file') return;
    const dir = vscode.Uri.joinPath(this.uri, '..');
    const base = this.uri.path.split('/').pop() ?? '';
    this.watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(dir, base));
    const onChange = () => {
      if (this.dirty || Date.now() - this.lastSave < 2500) return;
      this.run(async () => {
        const bytes = await vscode.workspace.fs.readFile(this.uri);
        await this.load(bytes, true);
        this.emit({ type: 'reset', meta: this.meta() });
      });
    };
    this.watcher.onDidChange(onChange);
    this.watcher.onDidCreate(onChange);
  }

  dispose(): void {
    this.watcher?.dispose();
    this._onDidDispose.fire();
    this._onDidEdit.dispose();
    this._onDidUpdate.dispose();
    this._onDidDispose.dispose();
  }

  // ------------------------------------------------------------------ loading

  private async load(bytes: Uint8Array, detectWarnings: boolean): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(bytes) as unknown as ExcelJS.Buffer);
    if (wb.worksheets.length === 0) wb.addWorksheet('Sheet1');
    unshareFormulas(wb);
    this.wb = wb;
    const themes = (wb as unknown as { _themes?: Record<string, string> })._themes;
    this.theme = parseTheme(themes?.theme1);
    this.rebuildHelpers();
    if (detectWarnings) this.warnings = await detectUnsupported(bytes);
    if (this.settings.recalc) {
      try {
        this.recalc.computeMissing();
      } catch (err) {
        console.error('[xlsx-editor] initial recalculation', err);
      }
    }
  }

  private rebuildHelpers(): void {
    const date1904 = !!(this.wb.properties as { date1904?: boolean } | undefined)?.date1904;
    this.serializer = new Serializer(this.theme, { locale: this.settings.locale, date1904 });
    this.recalc = new Recalc(this.wb, date1904);
  }

  updateSettings(settings: DocSettings): void {
    this.settings = settings;
    this.rebuildHelpers();
  }

  meta(): WorkbookMeta {
    return {
      fileName: this.uri.path.split('/').pop() ?? '',
      sheets: this.wb.worksheets.map((ws) => {
        const tab = resolveColor((ws.properties as { tabColor?: ExcelColor })?.tabColor, this.theme);
        return { name: ws.name, tab, hidden: ws.state !== undefined && ws.state !== 'visible' };
      }),
      warnings: this.warnings,
      defaultFont: this.theme.minorFont,
      readonly: false,
    };
  }

  sheetModel(index: number): SheetModel | undefined {
    const ws = this.wb.worksheets[index];
    if (!ws) return undefined;
    return this.serializer.sheet(ws, index);
  }

  // ------------------------------------------------------------------ queue

  /** Serializes all operations on the workbook. */
  run<T>(fn: () => T | Promise<T>): Promise<T | undefined> {
    const p = this.queue.then(fn).catch((err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      console.error('[xlsx-editor]', err);
      this.emit({ type: 'toast', message });
      return undefined;
    });
    this.queue = p;
    return p;
  }

  private emit(msg: ToWebview): void {
    this._onDidUpdate.fire(msg);
  }

  private fireEdit(label: string, undo: () => void | Promise<void>, redo: () => void | Promise<void>): void {
    this.dirty = true;
    this._onDidEdit.fire({
      label,
      undo: async () => {
        await this.run(undo);
        this.dirty = true;
      },
      redo: async () => {
        await this.run(redo);
        this.dirty = true;
      },
    });
  }

  private emitCells(keys: Iterable<string>): void {
    const bySheet = new Map<number, string[]>();
    for (const k of keys) {
      const s = Number(k.split(':')[0]);
      let list = bySheet.get(s);
      if (!list) bySheet.set(s, (list = []));
      list.push(k);
    }
    for (const [s, list] of bySheet) {
      const ws = this.wb.worksheets[s];
      if (!ws) continue;
      const cells: CellModel[] = [];
      const empty: [number, number][] = [];
      for (const k of list) {
        const [, r, c] = k.split(':').map(Number);
        const cell = ws.findCell(r, c);
        const m = cell ? this.serializer.cell(ws, cell, true) : undefined;
        if (m) cells.push(m);
        else empty.push([r, c]);
      }
      this.emit({ type: 'cells', sheet: s, cells, empty });
    }
  }

  private afterCellChange(changes: { s: number; r: number; c: number }[]): void {
    const keys = new Set(changes.map((ch) => cellKey(ch.s, ch.r, ch.c)));
    if (this.settings.recalc) {
      try {
        this.recalc.reindex(keys);
        for (const k of this.recalc.recalc(keys)) keys.add(k);
      } catch (err) {
        console.error('[xlsx-editor] recalculation', err);
      }
    }
    this.emitCells(keys);
  }

  private applySnaps(changes: CellChange[], which: 'before' | 'after'): void {
    for (const ch of changes) {
      const ws = this.wb.worksheets[ch.s];
      if (!ws) continue;
      const cell = ws.getCell(ch.r, ch.c);
      const s = ch[which];
      if (!s.skipValue && cell.type !== VT.Merge) cell.value = clone(s.value) as ExcelJS.CellValue;
      cell.style = clone(s.style) as Partial<ExcelJS.Style>;
    }
    this.afterCellChange(changes);
  }

  /** Cell-level operation, undoable via per-cell snapshots. */
  private cellOp(label: string, sheet: number, mutate: (touch: (r: number, c: number) => ExcelJS.Cell) => void): void {
    const ws = this.wb.worksheets[sheet];
    if (!ws) return;
    const touched = new Map<string, { r: number; c: number; before: CellSnap }>();
    const touch = (r: number, c: number) => {
      const k = cellKey(sheet, r, c);
      const cell = ws.getCell(r, c);
      if (!touched.has(k)) touched.set(k, { r, c, before: snap(cell) });
      return cell;
    };
    mutate(touch);
    const changes: CellChange[] = [];
    for (const t of touched.values()) {
      const after = snap(ws.getCell(t.r, t.c));
      if (JSON.stringify(after) !== JSON.stringify(t.before)) {
        changes.push({ s: sheet, r: t.r, c: t.c, before: t.before, after });
      }
    }
    if (!changes.length) return;
    this.afterCellChange(changes);
    this.fireEdit(
      label,
      () => this.applySnaps(changes, 'before'),
      () => this.applySnaps(changes, 'after'),
    );
  }

  /** Workbook-wide operation, undoable via a full workbook snapshot. */
  private async snapshotOp(label: string, mutate: () => number | void, recalcSheet?: number): Promise<void> {
    const before = new Uint8Array(await this.wb.xlsx.writeBuffer());
    let active: number | void;
    try {
      active = mutate();
    } catch (err) {
      await this.load(before, false);
      this.emit({ type: 'reset', meta: this.meta() });
      throw err;
    }
    this.rebuildHelpers();
    if (recalcSheet !== undefined && this.settings.recalc) {
      try {
        this.recalc.recalcSheet(recalcSheet);
      } catch (err) {
        console.error('[xlsx-editor] recalculation', err);
      }
    }
    const after = new Uint8Array(await this.wb.xlsx.writeBuffer());
    this.emit({ type: 'reset', meta: this.meta(), activeSheet: active ?? undefined });
    this.fireEdit(
      label,
      async () => {
        await this.load(before, false);
        this.emit({ type: 'reset', meta: this.meta() });
      },
      async () => {
        await this.load(after, false);
        this.emit({ type: 'reset', meta: this.meta(), activeSheet: active ?? undefined });
      },
    );
  }

  private clampRange(ws: ExcelJS.Worksheet, r: CellRange, existingOnly: boolean): CellRange {
    const maxR = existingOnly ? Math.max(ws.rowCount, 1) : r.r2;
    const maxC = existingOnly ? Math.max(ws.columnCount, 1) : r.c2;
    return { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, maxR), c2: Math.min(r.c2, maxC) };
  }

  private checkSize(ranges: CellRange[]): void {
    const n = ranges.reduce((acc, r) => acc + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0);
    if (n > MAX_CELLS_PER_OP) throw new Error(`Selection too large (${n.toLocaleString('en-US')} cells).`);
  }

  // ------------------------------------------------------------------ operations

  setCells(sheet: number, items: { r: number; c: number; input: string }[]): Promise<unknown> {
    return this.run(() =>
      this.cellOp('Edit cells', sheet, (touch) => {
        const ws = this.wb.worksheets[sheet];
        for (const it of items) {
          let cell = ws.getCell(it.r, it.c);
          if (cell.type === VT.Merge) cell = cell.master;
          cell = touch(Number(cell.row), Number(cell.col));
          this.writeInput(cell, it.input);
        }
      }),
    );
  }

  private writeInput(cell: ExcelJS.Cell, input: string): void {
    const parsed = parseInput(input, { locale: this.settings.locale, cellNumFmt: cell.numFmt });
    let value = parsed.value as unknown;
    const old = cell.value as unknown;
    if (old && typeof old === 'object' && typeof (old as { hyperlink?: unknown }).hyperlink === 'string' && typeof value === 'string') {
      value = { text: value, hyperlink: (old as { hyperlink: string }).hyperlink, tooltip: (old as { tooltip?: string }).tooltip };
    }
    cell.value = value as ExcelJS.CellValue;
    if (parsed.numFmt && (!cell.numFmt || cell.numFmt === 'General')) setStyle(cell, 'numFmt', parsed.numFmt);
  }

  clear(sheet: number, ranges: CellRange[], what: 'contents' | 'formats' | 'all'): Promise<unknown> {
    return this.run(() => {
      const ws = this.wb.worksheets[sheet];
      if (!ws) return;
      const clamped = ranges.map((r) => this.clampRange(ws, r, true));
      this.checkSize(clamped);
      this.cellOp(what === 'formats' ? 'Clear formats' : 'Clear', sheet, (touch) => {
        for (const rg of clamped) {
          for (let r = rg.r1; r <= rg.r2; r++) {
            for (let c = rg.c1; c <= rg.c2; c++) {
              const existing = ws.findCell(r, c);
              if (!existing) continue;
              const cell = touch(r, c);
              if (what !== 'formats' && cell.type !== VT.Merge) cell.value = null;
              if (what !== 'contents') cell.style = {};
            }
          }
        }
      });
    });
  }

  format(sheet: number, ranges: CellRange[], patch: FormatPatch): Promise<unknown> {
    return this.run(() => {
      const ws = this.wb.worksheets[sheet];
      if (!ws) return;
      this.checkSize(ranges);
      this.cellOp('Format', sheet, (touch) => {
        for (const rg of ranges) {
          for (let r = rg.r1; r <= rg.r2; r++) {
            for (let c = rg.c1; c <= rg.c2; c++) {
              const cell = touch(r, c);
              this.applyPatch(cell, patch, {
                top: r === rg.r1,
                bottom: r === rg.r2,
                left: c === rg.c1,
                right: c === rg.c2,
              });
            }
          }
          if (patch.border === 'none') {
            // also clear neighbouring borders that touch the range
            const clearSide = (r: number, c: number, side: keyof ExcelJS.Borders) => {
              if (r < 1 || c < 1 || !ws.findCell(r, c)) return;
              const cell = touch(r, c);
              const b = { ...(cell.style?.border ?? {}) } as Record<string, unknown>;
              if (!b[side]) return;
              delete b[side];
              setStyle(cell, 'border', b as Partial<ExcelJS.Borders>);
            };
            for (let c = rg.c1; c <= rg.c2; c++) {
              clearSide(rg.r1 - 1, c, 'bottom');
              clearSide(rg.r2 + 1, c, 'top');
            }
            for (let r = rg.r1; r <= rg.r2; r++) {
              clearSide(r, rg.c1 - 1, 'right');
              clearSide(r, rg.c2 + 1, 'left');
            }
          }
        }
      });
    });
  }

  private applyPatch(
    cell: ExcelJS.Cell,
    p: FormatPatch,
    edge: { top: boolean; bottom: boolean; left: boolean; right: boolean },
  ): void {
    const fontKeys: (keyof FormatPatch)[] = ['bold', 'italic', 'underline', 'strike', 'fontName', 'fontSize', 'fontColor'];
    if (fontKeys.some((k) => p[k] !== undefined)) {
      const font = { ...(cell.style?.font ?? {}) } as Partial<ExcelJS.Font>;
      if (p.bold !== undefined) font.bold = p.bold;
      if (p.italic !== undefined) font.italic = p.italic;
      if (p.underline !== undefined) font.underline = p.underline;
      if (p.strike !== undefined) font.strike = p.strike;
      if (p.fontName) font.name = p.fontName;
      if (p.fontSize) font.size = p.fontSize;
      if (p.fontColor === null) delete font.color;
      else if (p.fontColor) font.color = { argb: cssToArgb(p.fontColor) };
      setStyle(cell, 'font', font);
    }
    if (p.fill !== undefined) {
      setStyle(
        cell,
        'fill',
        p.fill === null
          ? { type: 'pattern', pattern: 'none' }
          : { type: 'pattern', pattern: 'solid', fgColor: { argb: cssToArgb(p.fill) }, bgColor: { indexed: 64 } as never },
      );
    }
    if (p.hAlign !== undefined || p.vAlign !== undefined || p.wrap !== undefined) {
      const al = { ...(cell.style?.alignment ?? {}) } as Record<string, unknown>;
      if (p.hAlign !== undefined) {
        if (p.hAlign === null) delete al.horizontal;
        else al.horizontal = p.hAlign;
      }
      if (p.vAlign !== undefined) {
        if (p.vAlign === null) delete al.vertical;
        else al.vertical = p.vAlign;
      }
      if (p.wrap !== undefined) al.wrapText = p.wrap;
      setStyle(cell, 'alignment', al as Partial<ExcelJS.Alignment>);
    }
    if (p.numFmt !== undefined) setStyle(cell, 'numFmt', p.numFmt);
    if (p.border) {
      const b = { ...(cell.style?.border ?? {}) } as Partial<ExcelJS.Borders>;
      switch (p.border) {
        case 'all':
          b.top = THIN;
          b.bottom = THIN;
          b.left = THIN;
          b.right = THIN;
          break;
        case 'none':
          delete b.top;
          delete b.bottom;
          delete b.left;
          delete b.right;
          break;
        case 'outer':
        case 'thickOuter': {
          const line = p.border === 'outer' ? THIN : MEDIUM;
          if (edge.top) b.top = line;
          if (edge.bottom) b.bottom = line;
          if (edge.left) b.left = line;
          if (edge.right) b.right = line;
          break;
        }
        case 'top':
          if (edge.top) b.top = THIN;
          break;
        case 'bottom':
          if (edge.bottom) b.bottom = THIN;
          break;
        case 'left':
          if (edge.left) b.left = THIN;
          break;
        case 'right':
          if (edge.right) b.right = THIN;
          break;
      }
      setStyle(cell, 'border', b);
    }
  }

  paste(
    sheet: number,
    target: CellRange,
    rows: string[][] | undefined,
    internal: { sheet: number; range: CellRange; cut: boolean } | undefined,
  ): Promise<unknown> {
    return this.run(() => {
      const ws = this.wb.worksheets[sheet];
      if (!ws) return;
      if (internal) return this.pasteInternal(sheet, target, internal);
      if (!rows || !rows.length) return;
      const h = rows.length;
      const w = Math.max(...rows.map((r) => r.length));
      const th = target.r2 - target.r1 + 1;
      const tw = target.c2 - target.c1 + 1;
      const repR = th % h === 0 ? th / h : 1;
      const repC = tw % w === 0 ? tw / w : 1;
      this.checkSize([{ r1: 1, c1: 1, r2: h * repR, c2: w * repC }]);
      this.cellOp('Paste', sheet, (touch) => {
        for (let rr = 0; rr < repR; rr++) {
          for (let cc = 0; cc < repC; cc++) {
            for (let i = 0; i < h; i++) {
              for (let j = 0; j < w; j++) {
                const r = target.r1 + rr * h + i;
                const c = target.c1 + cc * w + j;
                const existing = ws.getCell(r, c);
                if (existing.type === VT.Merge) continue;
                this.writeInput(touch(r, c), rows[i][j] ?? '');
              }
            }
          }
        }
      });
    });
  }

  private pasteInternal(
    sheet: number,
    target: CellRange,
    src: { sheet: number; range: CellRange; cut: boolean },
  ): Promise<void> | void {
    const sws = this.wb.worksheets[src.sheet];
    const tws = this.wb.worksheets[sheet];
    if (!sws || !tws) return;
    const { range } = src;
    const h = range.r2 - range.r1 + 1;
    const w = range.c2 - range.c1 + 1;
    const th = target.r2 - target.r1 + 1;
    const tw = target.c2 - target.c1 + 1;
    const repR = !src.cut && th % h === 0 ? th / h : 1;
    const repC = !src.cut && tw % w === 0 ? tw / w : 1;
    this.checkSize([{ r1: 1, c1: 1, r2: h * repR, c2: w * repC }]);
    // snapshot the source before any modification
    const source: { value: unknown; style: unknown; isFormula: boolean; formula?: string; merged: boolean }[][] = [];
    for (let i = 0; i < h; i++) {
      const row: (typeof source)[number] = [];
      for (let j = 0; j < w; j++) {
        const cell = sws.findCell(range.r1 + i, range.c1 + j);
        row.push({
          value: cell && cell.type !== VT.Merge ? clone(cell.value) : null,
          style: clone(cell?.style ?? {}),
          isFormula: cell?.type === VT.Formula,
          formula: cell?.type === VT.Formula ? cell.formula : undefined,
          merged: cell?.type === VT.Merge,
        });
      }
      source.push(row);
    }
    const label = src.cut ? 'Cut and paste' : 'Paste';
    const doPaste = (touch: (r: number, c: number) => ExcelJS.Cell) => {
      for (let rr = 0; rr < repR; rr++) {
        for (let cc = 0; cc < repC; cc++) {
          for (let i = 0; i < h; i++) {
            for (let j = 0; j < w; j++) {
              const r = target.r1 + rr * h + i;
              const c = target.c1 + cc * w + j;
              const s = source[i][j];
              const cell = touch(r, c);
              if (cell.type !== VT.Merge) {
                if (s.isFormula && s.formula) {
                  const dr = r - (range.r1 + i);
                  const dc = c - (range.c1 + j);
                  const f = src.cut ? s.formula : transformRefs(s.formula, (ref) => offsetRef(ref, dr, dc));
                  cell.value = { formula: f } as ExcelJS.CellValue;
                } else {
                  cell.value = clone(s.value) as ExcelJS.CellValue;
                }
              }
              cell.style = clone(s.style) as Partial<ExcelJS.Style>;
            }
          }
        }
      }
    };
    if (src.cut && src.sheet !== sheet) {
      // cut to another sheet: two sheets touched -> full snapshot
      return this.snapshotOp(label, () => {
        doPaste((r, c) => tws.getCell(r, c));
        this.clearSource(sws, range, target, false);
        return sheet;
      });
    }
    this.cellOp(label, sheet, (touch) => {
      doPaste(touch);
      if (src.cut) this.clearSource(sws, range, target, true, touch);
    });
  }

  private clearSource(
    ws: ExcelJS.Worksheet,
    range: CellRange,
    target: CellRange,
    sameSheet: boolean,
    touch?: (r: number, c: number) => ExcelJS.Cell,
  ): void {
    const h = range.r2 - range.r1 + 1;
    const w = range.c2 - range.c1 + 1;
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const inTarget = sameSheet && r >= target.r1 && r < target.r1 + h && c >= target.c1 && c < target.c1 + w;
        if (inTarget || !ws.findCell(r, c)) continue;
        const cell = touch ? touch(r, c) : ws.getCell(r, c);
        if (cell.type !== VT.Merge) cell.value = null;
        cell.style = {};
      }
    }
  }

  structure(sheet: number, op: StructureOp, at: number, count: number): Promise<unknown> {
    const labels: Record<StructureOp, string> = {
      insertRows: 'Insert rows',
      deleteRows: 'Delete rows',
      insertCols: 'Insert columns',
      deleteCols: 'Delete columns',
    };
    return this.run(() =>
      this.snapshotOp(
        labels[op],
        () => {
          const axis = op.endsWith('Rows') ? 'row' : 'col';
          spliceSheet(this.wb, sheet, axis, at, op.startsWith('insert') ? count : -count);
          return sheet;
        },
        sheet,
      ),
    );
  }

  merge(sheet: number, range: CellRange, unmerge: boolean): Promise<unknown> {
    return this.run(() =>
      this.snapshotOp(unmerge ? 'Unmerge cells' : 'Merge cells', () => {
        const ws = this.wb.worksheets[sheet];
        const merges = (ws as unknown as { _merges: Record<string, { model: { top: number; left: number; bottom: number; right: number } }> })._merges;
        for (const key of Object.keys(merges)) {
          const m = merges[key].model;
          const intersects = !(m.right < range.c1 || m.left > range.c2 || m.bottom < range.r1 || m.top > range.r2);
          if (intersects) ws.unMergeCells(rangeToA1(m.top, m.left, m.bottom, m.right));
        }
        if (!unmerge && (range.r1 !== range.r2 || range.c1 !== range.c2)) {
          ws.mergeCells(range.r1, range.c1, range.r2, range.c2);
        }
        return sheet;
      }),
    );
  }

  colWidth(sheet: number, cols: number[], px: number): Promise<unknown> {
    return this.run(() => {
      const ws = this.wb.worksheets[sheet];
      if (!ws) return;
      const before = cols.map((c) => ws.getColumn(c).width);
      const width = pxToColWidth(px);
      const apply = (widths: (number | undefined)[]) => {
        cols.forEach((c, i) => {
          ws.getColumn(c).width = widths[i] as number;
        });
        this.emitSheet(sheet);
      };
      apply(cols.map(() => width));
      this.fireEdit('Column width', () => apply(before), () => apply(cols.map(() => width)));
    });
  }

  rowHeight(sheet: number, rows: number[], px: number): Promise<unknown> {
    return this.run(() => {
      const ws = this.wb.worksheets[sheet];
      if (!ws) return;
      const before = rows.map((r) => ws.getRow(r).height);
      const height = pxToPt(px);
      const apply = (heights: (number | undefined)[]) => {
        rows.forEach((r, i) => {
          ws.getRow(r).height = heights[i] as number;
        });
        this.emitSheet(sheet);
      };
      apply(rows.map(() => height));
      this.fireEdit('Row height', () => apply(before), () => apply(rows.map(() => height)));
    });
  }

  private emitSheet(index: number): void {
    const data = this.sheetModel(index);
    if (data) this.emit({ type: 'sheet', index, data });
  }

  addSheet(): Promise<unknown> {
    return this.run(() =>
      this.snapshotOp('New sheet', () => {
        let n = this.wb.worksheets.length + 1;
        while (this.wb.worksheets.some((ws) => ws.name.toLowerCase() === `sheet${n}`)) n++;
        this.wb.addWorksheet(`Sheet${n}`);
        return this.wb.worksheets.length - 1;
      }),
    );
  }

  renameSheet(index: number, name: string): Promise<unknown> {
    return this.run(async () => {
      const ws = this.wb.worksheets[index];
      if (!ws || ws.name === name.trim()) return;
      const err = validateSheetName(this.wb, name, index);
      if (err) throw new Error(err);
      await this.snapshotOp('Rename sheet', () => {
        renameSheet(this.wb, index, name.trim());
        return index;
      });
    });
  }

  deleteSheet(index: number): Promise<unknown> {
    return this.run(async () => {
      const visible = this.wb.worksheets.filter((ws) => ws.state === 'visible' || !ws.state);
      const ws = this.wb.worksheets[index];
      if (!ws) return;
      if (visible.length <= 1 && (ws.state === 'visible' || !ws.state)) {
        throw new Error('A workbook must contain at least one visible sheet.');
      }
      await this.snapshotOp('Delete sheet', () => {
        deleteSheet(this.wb, index);
        return Math.max(0, index - 1);
      });
    });
  }

  /** Values offered by a "list" data validation. */
  listValues(sheet: number, r: number, c: number): string[] {
    const ws = this.wb.worksheets[sheet];
    if (!ws) return [];
    const dv = dataValidations(ws).find(address(r, c)) as { type?: string; formulae?: unknown[] } | undefined;
    if (!dv || dv.type !== 'list' || !dv.formulae?.length) return [];
    const f = String(dv.formulae[0] ?? '').trim();
    if (f.startsWith('"')) {
      return f
        .slice(1, f.endsWith('"') ? -1 : undefined)
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    }
    let text = f.replace(/^=/, '');
    let ref = parseRangeText(text);
    if (!ref) {
      try {
        const ranges = (this.wb.definedNames as unknown as { getRanges(n: string): { ranges: string[] } }).getRanges(text)?.ranges;
        if (ranges?.length) {
          text = ranges[0];
          ref = parseRangeText(text);
        }
      } catch {
        /* ignored */
      }
    }
    if (!ref) return [];
    const target = ref.sheet ? this.wb.worksheets.find((w) => w.name.toLowerCase() === ref!.sheet!.toLowerCase()) : ws;
    if (!target) return [];
    const out: string[] = [];
    const r2 = Math.min(ref.r2, target.rowCount);
    const c2 = Math.min(ref.c2, Math.max(target.columnCount, ref.c1));
    for (let rr = ref.r1; rr <= r2 && out.length < 1000; rr++) {
      for (let cc = ref.c1; cc <= c2; cc++) {
        const cell = target.findCell(rr, cc);
        if (!cell) continue;
        const m = this.serializer.cell(target, cell, false);
        const t = m?.t ?? richTextToString(cell.value);
        if (t && !out.includes(t)) out.push(t);
      }
    }
    return out;
  }

  // ------------------------------------------------------------------ persistence

  private async bytes(): Promise<Uint8Array> {
    const calc = (this.wb as unknown as { calcProperties: Record<string, unknown> }).calcProperties ?? {};
    (this.wb as unknown as { calcProperties: Record<string, unknown> }).calcProperties = { ...calc, fullCalcOnLoad: true };
    return new Uint8Array(await this.wb.xlsx.writeBuffer());
  }

  save(_cancel: vscode.CancellationToken): Promise<void> {
    return this.saveAs(this.uri, _cancel).then(() => {
      this.dirty = false;
    });
  }

  async saveAs(target: vscode.Uri, _cancel: vscode.CancellationToken): Promise<void> {
    await this.queue.catch(() => undefined);
    const bytes = await this.bytes();
    this.lastSave = Date.now();
    await vscode.workspace.fs.writeFile(target, bytes);
    this.lastSave = Date.now();
  }

  async revert(_cancel: vscode.CancellationToken): Promise<void> {
    await this.run(async () => {
      const bytes = await vscode.workspace.fs.readFile(this.uri);
      await this.load(bytes, true);
      this.dirty = false;
      this.emit({ type: 'reset', meta: this.meta() });
    });
  }

  async backup(destination: vscode.Uri, cancel: vscode.CancellationToken): Promise<vscode.CustomDocumentBackup> {
    await this.queue.catch(() => undefined);
    if (!cancel.isCancellationRequested) {
      const bytes = new Uint8Array(await this.wb.xlsx.writeBuffer());
      await vscode.workspace.fs.writeFile(destination, bytes);
    }
    return {
      id: destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(destination);
        } catch {
          /* already deleted */
        }
      },
    };
  }
}

/** Detects file features that ExcelJS cannot write back. */
async function detectUnsupported(bytes: Uint8Array): Promise<string[]> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files);
    const has = (re: RegExp) => names.some((n) => re.test(n));
    const out: string[] = [];
    if (has(/^xl\/charts\//)) out.push('charts');
    if (has(/^xl\/pivotTables\//) || has(/^xl\/pivotCache\//)) out.push('pivot tables');
    if (has(/^xl\/vbaProject\.bin$/)) out.push('VBA macros');
    if (has(/^xl\/slicers\//) || has(/^xl\/slicerCaches\//)) out.push('slicers');
    if (has(/^xl\/timelines\//)) out.push('timelines');
    if (has(/^xl\/externalLinks\//)) out.push('external links');
    if (has(/^xl\/threadedComments\//)) out.push('threaded comments');
    if (has(/^xl\/(ctrlProps|activeX)\//)) out.push('form controls');
    // shapes / text boxes in drawings
    for (const n of names.filter((x) => /^xl\/drawings\/drawing\d+\.xml$/.test(x))) {
      const xml = await zip.file(n)!.async('string');
      if (/<xdr:sp\b/.test(xml)) {
        out.push('shapes and text boxes');
        break;
      }
    }
    return out;
  } catch {
    return [];
  }
}
