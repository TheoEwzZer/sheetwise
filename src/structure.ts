// Row/column insertion and deletion, sheet renaming and deletion,
// rewriting formulas, merges, data validations and conditional formatting.
import type ExcelJS from 'exceljs';
import {
  type Axis,
  type Ref,
  address,
  colToNum,
  parseRangeText,
  refToString,
  spliceRef,
  transformRefs,
  MAX_COL,
  MAX_ROW,
} from './refs.ts';
import { internalCells, internalRows, VT, dataValidations } from './serialize.ts';

type Worksheet = ExcelJS.Worksheet;

const sameName = (a: string | undefined, b: string) => a !== undefined && a.toLowerCase() === b.toLowerCase();

function decodeAddress(a: string): [number, number] | undefined {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(a);
  if (!m) return undefined;
  return [+m[2], colToNum(m[1])];
}

/** Applies `fn` to every formula in the workbook. */
export function rewriteAllFormulas(
  wb: ExcelJS.Workbook,
  fn: (ref: Ref, ownerSheet: string) => Ref | null | undefined,
  skipSheet?: number,
): void {
  wb.worksheets.forEach((ws, si) => {
    if (si === skipSheet) return;
    internalRows(ws).forEach((row) => {
      if (!row) return;
      internalCells(row).forEach((cell) => {
        if (!cell || cell.type !== VT.Formula) return;
        const v = cell.value as unknown as Record<string, unknown>;
        const f = cell.formula;
        if (!f) return;
        const nf = transformRefs(f, (ref) => fn(ref, ws.name));
        if (nf !== f) {
          const copy: Record<string, unknown> = { ...v, formula: nf, result: cell.result };
          delete copy.sharedFormula;
          cell.value = copy as unknown as ExcelJS.CellValue;
        }
      });
    });
  });
}

/** Converts all shared formulas into explicit formulas (makes edits safe). */
export function unshareFormulas(wb: ExcelJS.Workbook): void {
  for (const ws of wb.worksheets) {
    const updates: [ExcelJS.Cell, Record<string, unknown>][] = [];
    internalRows(ws).forEach((row) => {
      if (!row) return;
      internalCells(row).forEach((cell) => {
        if (!cell || cell.type !== VT.Formula) return;
        const v = cell.value as unknown as Record<string, unknown>;
        if (v.sharedFormula !== undefined || v.shareType === 'shared') {
          const formula = cell.formula;
          if (formula) updates.push([cell, { formula, result: cell.result }]);
        }
      });
    });
    // compute all translated formulas before modifying the masters
    for (const [cell, v] of updates) cell.value = v as unknown as ExcelJS.CellValue;
  }
}

function transformRangeList(text: string, fn: (ref: Ref) => Ref | null | undefined): string | null {
  const parts = text.split(/\s+/).filter(Boolean);
  const out: string[] = [];
  for (const p of parts) {
    const ref = parseRangeText(p);
    if (!ref) {
      out.push(p);
      continue;
    }
    const res = fn(ref);
    if (res === null) continue;
    out.push(refToString(res ?? ref));
  }
  return out.length ? out.join(' ') : null;
}

function spliceIndex(n: number, at: number, count: number): number | null {
  if (count > 0) return n >= at ? n + count : n;
  const del = -count;
  if (n >= at && n < at + del) return null;
  return n >= at + del ? n - del : n;
}

interface TableLike {
  name?: string;
  ref?: string;
  tableRef?: string;
  autoFilterRef?: string;
}

export function spliceSheet(wb: ExcelJS.Workbook, sheetIndex: number, axis: Axis, at: number, count: number): void {
  const ws = wb.worksheets[sheetIndex];
  const name = ws.name;
  const fnSelf = (ref: Ref) => (ref.sheet === undefined || sameName(ref.sheet, name) ? spliceRef(ref, axis, at, count) : undefined);

  const model = ws.model as unknown as Record<string, any>;

  // Tables (Excel Table objects) do not support inserting/deleting columns inside them.
  const tables = (model.tables ?? []) as TableLike[];
  for (const t of tables) {
    const ref = t.ref ? parseRangeText(t.ref) : undefined;
    if (!ref) continue;
    if (axis === 'col') {
      const lo = Math.min(ref.c1, ref.c2);
      const hi = Math.max(ref.c1, ref.c2);
      const touches = count > 0 ? at > lo && at <= hi : at <= hi && at - count - 1 >= lo;
      if (touches) {
        throw new Error(`Cannot change the columns of table "${t.name ?? ''}": use Excel for this operation.`);
      }
    } else if (count < 0) {
      const lo = Math.min(ref.r1, ref.r2);
      if (at <= lo && at - count - 1 >= lo) {
        throw new Error(`Cannot delete the header row of table "${t.name ?? ''}".`);
      }
    }
  }

  // 1) formulas on other sheets that point to this one
  rewriteAllFormulas(
    wb,
    (ref) => (sameName(ref.sheet, name) ? spliceRef(ref, axis, at, count) : undefined),
    sheetIndex,
  );

  // 2) rows and cells
  const newRows: any[] = [];
  for (const rowModel of model.rows as any[]) {
    let number = rowModel.number as number;
    if (axis === 'row') {
      const n = spliceIndex(number, at, count);
      if (n === null || n > MAX_ROW) continue;
      number = n;
    }
    const cells: any[] = [];
    let min = 0;
    let max = 0;
    for (const cm of rowModel.cells as any[]) {
      const pos = decodeAddress(cm.address);
      if (!pos) continue;
      let [, c] = pos;
      if (axis === 'col') {
        const n = spliceIndex(c, at, count);
        if (n === null || n > MAX_COL) continue;
        c = n;
      }
      const next: any = { ...cm, address: address(number, c) };
      if (next.type === VT.Merge) {
        next.type = VT.Null;
        delete next.master;
        delete next.value;
      }
      if (typeof next.formula === 'string') next.formula = transformRefs(next.formula, fnSelf);
      if (typeof next.ref === 'string') next.ref = transformRangeList(next.ref, fnSelf) ?? next.address;
      delete next.sharedFormula;
      cells.push(next);
      min = min ? Math.min(min, c) : c;
      max = Math.max(max, c);
    }
    newRows.push({ ...rowModel, number, cells, min: min || rowModel.min, max: max || rowModel.max });
  }
  newRows.sort((a, b) => a.number - b.number);
  model.rows = newRows;

  // 3) merges
  const merges: string[] = [];
  for (const m of (model.merges ?? []) as string[]) {
    const res = transformRangeList(m, fnSelf);
    if (res && res.includes(':')) {
      const ref = parseRangeText(res);
      if (ref && (ref.r1 !== ref.r2 || ref.c1 !== ref.c2)) merges.push(res);
    }
  }
  model.mergeCells = merges;

  // 4) column definitions
  if (axis === 'col' && Array.isArray(model.cols)) {
    const cols: any[] = [];
    for (const col of model.cols) {
      const ref: Ref = { kind: 'cols', r1: 1, r2: MAX_ROW, c1: col.min, c2: col.max, ar1: false, ar2: false, ac1: false, ac2: false };
      const res = spliceRef(ref, 'col', at, count);
      if (res === null) continue;
      const r = res ?? ref;
      cols.push({ ...col, min: r.c1, max: r.c2 });
    }
    model.cols = cols;
  }

  // 5) data validations
  const dvs = (model.dataValidations ?? {}) as Record<string, any>;
  const newDvs: Record<string, any> = {};
  for (const key of Object.keys(dvs)) {
    const dv = dvs[key];
    if (!dv) continue;
    const nk = transformRangeList(key, fnSelf);
    if (!nk) continue;
    const copy = { ...dv };
    if (Array.isArray(copy.formulae)) {
      copy.formulae = copy.formulae.map((f: unknown) => (typeof f === 'string' ? transformRefs(f, fnSelf) : f));
    }
    newDvs[nk] = copy;
  }
  // Like Excel: inserting inside a validated range extends the validation.
  if (count > 0 && at > 1) {
    for (const key of Object.keys(newDvs)) {
      const pos = decodeAddress(key);
      if (!pos) continue;
      const [r, c] = pos;
      const before = axis === 'row' ? r === at - 1 : c === at - 1;
      if (!before) continue;
      const afterKey = axis === 'row' ? address(at + count, c) : address(r, at + count);
      if (!newDvs[afterKey] || JSON.stringify(newDvs[afterKey]) !== JSON.stringify(newDvs[key])) continue;
      for (let i = 0; i < count; i++) {
        const k = axis === 'row' ? address(at + i, c) : address(r, at + i);
        if (!newDvs[k]) newDvs[k] = { ...newDvs[key] };
      }
    }
  }
  model.dataValidations = newDvs;

  // 6) conditional formatting
  if (Array.isArray(model.conditionalFormattings)) {
    model.conditionalFormattings = model.conditionalFormattings
      .map((cf: any) => {
        const ref = typeof cf.ref === 'string' ? transformRangeList(cf.ref, fnSelf) : cf.ref;
        if (!ref) return null;
        const rules = (cf.rules ?? []).map((rule: any) =>
          Array.isArray(rule.formulae)
            ? { ...rule, formulae: rule.formulae.map((f: unknown) => (typeof f === 'string' ? transformRefs(f, fnSelf) : f)) }
            : rule,
        );
        return { ...cf, ref, rules };
      })
      .filter(Boolean);
  }

  // 7) autofilter
  if (typeof model.autoFilter === 'string') {
    model.autoFilter = transformRangeList(model.autoFilter, fnSelf) ?? undefined;
  }

  // 8) tables
  model.tables = tables.map((t) => {
    const copy: any = { ...t };
    for (const k of ['ref', 'tableRef', 'autoFilterRef'] as const) {
      if (typeof copy[k] === 'string') copy[k] = transformRangeList(copy[k], fnSelf) ?? copy[k];
    }
    return copy;
  });

  // 9) images (0-based anchors)
  if (Array.isArray(model.media)) {
    const key = axis === 'row' ? 'nativeRow' : 'nativeCol';
    for (const medium of model.media) {
      for (const corner of ['tl', 'br'] as const) {
        const anchor = medium?.range?.[corner];
        if (anchor && typeof anchor[key] === 'number') {
          const idx = anchor[key] + 1;
          const n = spliceIndex(idx, at, count);
          anchor[key] = (n === null ? at : n) - 1;
        }
      }
    }
  }

  // 10) defined names
  const dn = wb.definedNames as unknown as {
    spliceRows(sheet: string, start: number, del: number, ins: number): void;
    spliceColumns(sheet: string, start: number, del: number, ins: number): void;
  };
  try {
    if (axis === 'row') dn.spliceRows(name, at, count < 0 ? -count : 0, count > 0 ? count : 0);
    else dn.spliceColumns(name, at, count < 0 ? -count : 0, count > 0 ? count : 0);
  } catch {
    /* ignored */
  }

  // Re-inject the model
  (ws as unknown as { _merges: object })._merges = {};
  ws.model = model as unknown as ExcelJS.WorksheetModel;

  // 11) formulas of the sheet itself that are outside the cell model (none) — nothing to do.

  // 12) inserted rows/columns inherit the format of the row above / column to the left
  if (count > 0 && at > 1) {
    if (axis === 'row') {
      const src = ws.findRow(at - 1);
      if (src) {
        internalCells(src).forEach((cell) => {
          if (!cell) return;
          for (let r = at; r < at + count; r++) ws.getCell(r, Number(cell.col)).style = { ...cell.style };
        });
      }
    } else {
      internalRows(ws).forEach((row) => {
        if (!row) return;
        const cell = row.findCell(at - 1);
        if (!cell) return;
        for (let c = at; c < at + count; c++) row.getCell(c).style = { ...cell.style };
      });
    }
  }
}

const INVALID_SHEET = /[\\/?*[\]:]/;

export function validateSheetName(wb: ExcelJS.Workbook, name: string, exceptIndex?: number): string | undefined {
  const n = name.trim();
  if (!n) return 'The sheet name cannot be empty.';
  if (n.length > 31) return 'The sheet name cannot exceed 31 characters.';
  if (INVALID_SHEET.test(n)) return 'The sheet name cannot contain: \\ / ? * [ ] :';
  if (n.startsWith("'") || n.endsWith("'")) return "The sheet name cannot start or end with an apostrophe.";
  if (wb.worksheets.some((ws, i) => i !== exceptIndex && ws.name.toLowerCase() === n.toLowerCase())) {
    return 'A sheet with this name already exists.';
  }
  return undefined;
}

export function renameSheet(wb: ExcelJS.Workbook, index: number, newName: string): void {
  const ws = wb.worksheets[index];
  const old = ws.name;
  rewriteAllFormulas(wb, (ref) => (sameName(ref.sheet, old) ? { ...ref, sheet: newName } : undefined));
  // data validations on other sheets
  for (const other of wb.worksheets) {
    const dvs = dataValidations(other).model;
    for (const k of Object.keys(dvs)) {
      const dv = dvs[k];
      if (dv && Array.isArray(dv.formulae)) {
        dv.formulae = dv.formulae.map((f: unknown) =>
          typeof f === 'string' ? transformRefs(f, (ref) => (sameName(ref.sheet, old) ? { ...ref, sheet: newName } : undefined)) : f,
        );
      }
    }
  }
  ws.name = newName;
}

export function deleteSheet(wb: ExcelJS.Workbook, index: number): void {
  const ws = wb.worksheets[index];
  const name = ws.name;
  rewriteAllFormulas(wb, (ref) => (sameName(ref.sheet, name) ? null : undefined), index);
  wb.removeWorksheet(ws.id);
}
