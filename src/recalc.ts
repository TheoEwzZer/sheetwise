// Recalculates formulas that depend on modified cells (fast-formula-parser engine).
import type ExcelJS from 'exceljs';
import FormulaParser, { type FormulaError } from 'fast-formula-parser';
import { findRefs, transformRefs, MAX_ROW, MAX_COL, parseRangeText } from './refs.ts';
import { dateToSerial } from './numfmt.ts';
import { EXTRA_FUNCTIONS, patchEmptyComparison } from './formulaFunctions.ts';

type Cell = ExcelJS.Cell;
type Worksheet = ExcelJS.Worksheet;

const VT_MERGE = 1;
const VT_FORMULA = 6;

interface RangeDep {
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  f: string;
}

export const cellKey = (sheet: number, r: number, c: number) => `${sheet}:${r}:${c}`;

function parseKey(k: string): [number, number, number] {
  const [s, r, c] = k.split(':').map(Number);
  return [s, r, c];
}

/** A formula result as stored by ExcelJS. */
type ResultValue = number | string | boolean | Date | { error: string } | undefined;

export class Recalc {
  private indexed = false;
  private formulaDeps = new Map<string, { cells: string[]; ranges: RangeDep[] }>();
  private reverse = new Map<string, Set<string>>();
  private ranges: RangeDep[] = [];
  private evalText = new Map<string, string>();
  private parser: FormulaParser;

  // state of a single evaluation pass
  private affected = new Set<string>();
  private memo = new Map<string, unknown>();
  private visiting = new Set<string>();

  private wb: ExcelJS.Workbook;
  private date1904: boolean;

  constructor(wb: ExcelJS.Workbook, date1904: boolean) {
    this.wb = wb;
    this.date1904 = date1904;
    patchEmptyComparison();
    this.parser = new FormulaParser({
      functions: EXTRA_FUNCTIONS as unknown as Record<string, (...args: unknown[]) => unknown>,
      onCell: ({ sheet, row, col }) => this.valueAt(this.sheetIndex(sheet), row, col),
      onRange: (ref) => this.rangeValues(this.sheetIndex(ref.sheet), ref.from.row, ref.from.col, ref.to.row, ref.to.col),
      onVariable: (name, sheetName) => this.resolveName(name, sheetName),
    });
  }

  private sheets(): Worksheet[] {
    return this.wb.worksheets;
  }

  sheetIndex(name: string | undefined): number {
    if (name === undefined) return -1;
    const lower = name.toLowerCase();
    return this.sheets().findIndex((s) => s.name.toLowerCase() === lower);
  }

  invalidate(): void {
    this.indexed = false;
    this.formulaDeps.clear();
    this.reverse.clear();
    this.ranges = [];
    this.evalText.clear();
  }

  private forEachFormula(cb: (sheet: number, cell: Cell) => void): void {
    this.sheets().forEach((ws, si) => {
      const rows = (ws as unknown as { _rows: (ExcelJS.Row | undefined)[] })._rows;
      rows.forEach((row) => {
        if (!row) return;
        const cells = (row as unknown as { _cells: (Cell | undefined)[] })._cells;
        cells.forEach((cell) => {
          if (cell && cell.type === VT_FORMULA) cb(si, cell);
        });
      });
    });
  }

  private ensureIndex(): void {
    if (this.indexed) return;
    this.forEachFormula((si, cell) => this.indexCell(si, cell));
    this.indexed = true;
  }

  private removeIndex(key: string): void {
    const deps = this.formulaDeps.get(key);
    if (!deps) return;
    for (const c of deps.cells) this.reverse.get(c)?.delete(key);
    if (deps.ranges.length) this.ranges = this.ranges.filter((r) => r.f !== key);
    this.formulaDeps.delete(key);
    this.evalText.delete(key);
  }

  private indexCell(si: number, cell: Cell): void {
    const key = cellKey(si, Number(cell.row), Number(cell.col));
    this.removeIndex(key);
    const formula = cell.formula;
    if (!formula) return;
    const deps = { cells: [] as string[], ranges: [] as RangeDep[] };
    for (const ref of findRefs(formula)) {
      const s = ref.sheet !== undefined ? this.sheetIndex(ref.sheet) : si;
      if (s < 0) continue;
      if (ref.kind === 'cell') {
        const k = cellKey(s, ref.r1, ref.c1);
        deps.cells.push(k);
        let set = this.reverse.get(k);
        if (!set) this.reverse.set(k, (set = new Set()));
        set.add(key);
      } else {
        const rd: RangeDep = {
          sheet: s,
          r1: Math.min(ref.r1, ref.r2),
          c1: Math.min(ref.c1, ref.c2),
          r2: Math.max(ref.r1, ref.r2),
          c2: Math.max(ref.c1, ref.c2),
          f: key,
        };
        deps.ranges.push(rd);
        this.ranges.push(rd);
      }
    }
    // Defined names are not tracked precisely: treat them as depending on everything.
    if (/[A-Za-z_]\w*/.test(formula) && this.wb.definedNames && this.mentionsDefinedName(formula)) {
      const rd: RangeDep = { sheet: -1, r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL, f: key };
      deps.ranges.push(rd);
      this.ranges.push(rd);
    }
    this.formulaDeps.set(key, deps);
  }

  private definedNameList: string[] | undefined;
  private mentionsDefinedName(formula: string): boolean {
    if (!this.definedNameList) {
      try {
        const model = (this.wb.definedNames as unknown as { model: { name: string }[] }).model ?? [];
        this.definedNameList = model.map((m) => m.name.toLowerCase());
      } catch {
        this.definedNameList = [];
      }
    }
    if (!this.definedNameList.length) return false;
    const lower = formula.toLowerCase();
    return this.definedNameList.some((n) => new RegExp(`(^|[^\\w.])${n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^\\w(])`).test(lower));
  }

  /** Call after cells are modified (updates the dependency index). */
  reindex(keys: Iterable<string>): void {
    if (!this.indexed) return;
    for (const k of keys) {
      const [s, r, c] = parseKey(k);
      const ws = this.sheets()[s];
      if (!ws) continue;
      const cell = ws.findCell(r, c);
      if (cell && cell.type === VT_FORMULA) this.indexCell(s, cell);
      else this.removeIndex(k);
    }
  }

  /**
   * Recalculates the formulas affected by `changed`. Returns the keys whose result changed.
   */
  recalc(changed: Iterable<string>): string[] {
    this.ensureIndex();
    const affected = new Set<string>();
    const queue: string[] = [];
    for (const k of changed) {
      queue.push(k);
      if (this.formulaDeps.has(k)) affected.add(k);
    }
    while (queue.length) {
      const k = queue.pop()!;
      const [s, r, c] = parseKey(k);
      const direct = this.reverse.get(k);
      if (direct) {
        for (const f of direct) {
          if (!affected.has(f)) {
            affected.add(f);
            queue.push(f);
          }
        }
      }
      for (const rd of this.ranges) {
        if ((rd.sheet === s || rd.sheet === -1) && r >= rd.r1 && r <= rd.r2 && c >= rd.c1 && c <= rd.c2 && !affected.has(rd.f)) {
          affected.add(rd.f);
          queue.push(rd.f);
        }
      }
    }
    return this.evaluate(affected);
  }

  /** After a structural change: recalculates everything that touches sheet `s`. */
  recalcSheet(s: number): string[] {
    this.invalidate();
    this.ensureIndex();
    const seeds = new Set<string>();
    const prefix = `${s}:`;
    for (const [f, deps] of this.formulaDeps) {
      if (
        f.startsWith(prefix) ||
        deps.cells.some((k) => k.startsWith(prefix)) ||
        deps.ranges.some((r) => r.sheet === s || r.sheet === -1)
      ) {
        seeds.add(f);
      }
    }
    return this.recalc(seeds);
  }

  /** Computes formulas that have no cached result (e.g. script-generated files). */
  computeMissing(): string[] {
    const missing = new Set<string>();
    this.forEachFormula((si, cell) => {
      const v = cell.value as { result?: unknown };
      if (v && cell.result === undefined) missing.add(cellKey(si, Number(cell.row), Number(cell.col)));
    });
    if (!missing.size) return [];
    this.ensureIndex();
    // also includes the formulas that depend on these
    return this.recalc(missing);
  }

  private evaluate(affected: Set<string>): string[] {
    this.affected = affected;
    this.memo = new Map();
    this.visiting = new Set();
    const changed: string[] = [];
    for (const key of affected) {
      const [s, r, c] = parseKey(key);
      const cell = this.sheets()[s]?.findCell(r, c);
      if (!cell || cell.type !== VT_FORMULA) continue;
      const before = cell.result as ResultValue;
      const result = this.evalFormula(key, cell);
      if (result === KEEP) continue;
      if (!sameResult(before, result as ResultValue)) {
        const v = cell.value as unknown as Record<string, unknown>;
        cell.value = { ...v, result } as ExcelJS.CellValue;
        changed.push(key);
      }
    }
    this.affected = new Set();
    this.memo = new Map();
    return changed;
  }

  private evalFormula(key: string, cell: Cell): unknown {
    if (this.memo.has(key)) return this.memo.get(key);
    const cached = cell.result as ResultValue;
    if (this.visiting.has(key)) return cached === undefined ? 0 : cached; // circular reference
    this.visiting.add(key);
    let result: unknown = KEEP;
    try {
      const [s, r, c] = parseKey(key);
      const ws = this.sheets()[s];
      let text = this.evalText.get(key);
      if (text === undefined) {
        // fast-formula-parser mishandles "$A:$A": strip the "$" (no effect on the result)
        text = transformRefs(cell.formula, (ref) =>
          ref.ar1 || ref.ac1 || ref.ar2 || ref.ac2 ? { ...ref, ar1: false, ac1: false, ar2: false, ac2: false } : undefined,
        );
        this.evalText.set(key, text);
      }
      const raw = this.parser.parse(text, { sheet: ws.name, row: r, col: c }, true);
      result = toResult(raw);
    } catch (err) {
      if (err instanceof FormulaParser.FormulaError || (err as { _error?: string })._error) {
        const fe = err as FormulaError;
        const msg = String(fe.message ?? '');
        // function not supported by the engine: keep the value computed by Excel
        result = /not implemented/i.test(msg) || /#ERROR!/.test(msg) ? KEEP : { error: fe.error };
      } else {
        result = KEEP;
      }
    }
    this.visiting.delete(key);
    this.memo.set(key, result);
    return result;
  }

  private valueAt(s: number, r: number, c: number): unknown {
    const ws = this.sheets()[s];
    if (!ws) return new FormulaParser.FormulaError('#REF!');
    const cell = ws.findCell(r, c);
    if (!cell) return null;
    if (cell.type === VT_MERGE) return null;
    const key = cellKey(s, r, c);
    if (cell.type === VT_FORMULA) {
      if (this.affected.has(key)) {
        const res = this.evalFormula(key, cell);
        if (res !== KEEP) return toParserValue(res as ResultValue, this.date1904);
      }
      return toParserValue(cell.result, this.date1904);
    }
    return toParserValue(cell.value as ResultValue, this.date1904);
  }

  private rangeValues(s: number, r1: number, c1: number, r2: number, c2: number): unknown[][] {
    const ws = this.sheets()[s];
    if (!ws) return [[new FormulaParser.FormulaError('#REF!')]];
    const lastRow = Math.min(r2, Math.max(ws.rowCount, r1));
    const lastCol = Math.min(c2, Math.max(ws.columnCount, c1));
    const out: unknown[][] = [];
    for (let r = r1; r <= lastRow; r++) {
      const row: unknown[] = [];
      for (let c = c1; c <= lastCol; c++) row.push(this.valueAt(s, r, c));
      out.push(row);
    }
    return out.length ? out : [[null]];
  }

  private resolveName(name: string, sheetName: string) {
    try {
      const dn = this.wb.definedNames as unknown as { getRanges(n: string): { ranges: string[] } };
      const ranges = dn.getRanges(name)?.ranges;
      if (!ranges?.length) return undefined;
      const ref = parseRangeText(ranges[0]);
      if (!ref) return undefined;
      const sheet = ref.sheet ?? sheetName;
      if (ref.kind === 'cell') return { sheet, row: ref.r1, col: ref.c1 };
      return { sheet, from: { row: ref.r1, col: ref.c1 }, to: { row: ref.r2, col: ref.c2 } };
    } catch {
      return undefined;
    }
  }
}

const KEEP = Symbol('keep');

function toParserValue(v: ResultValue | unknown, date1904: boolean): unknown {
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return dateToSerial(v, date1904);
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if (typeof o.error === 'string') return new FormulaParser.FormulaError(o.error);
    if (Array.isArray(o.richText)) return (o.richText as { text: string }[]).map((t) => t.text).join('');
    if (o.text !== undefined) {
      const t = o.text as unknown;
      if (typeof t === 'string') return t;
      if (t && typeof t === 'object' && Array.isArray((t as { richText?: unknown }).richText)) {
        return ((t as { richText: { text: string }[] }).richText).map((x) => x.text).join('');
      }
    }
    if ('result' in o) return toParserValue(o.result, date1904);
    return null;
  }
  return v;
}

function toResult(raw: unknown): unknown {
  let v = raw;
  while (Array.isArray(v)) v = v[0];
  if (v === undefined || v === null) return 0;
  if (v instanceof FormulaParser.FormulaError || (v as { _error?: string })._error) {
    return { error: (v as FormulaError).error };
  }
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return { error: '#NUM!' };
    return v;
  }
  if (typeof v === 'string' || typeof v === 'boolean') return v;
  return KEEP;
}

function sameResult(a: ResultValue, b: ResultValue): boolean {
  if (a instanceof Date && typeof b === 'number') return false;
  if (a && b && typeof a === 'object' && typeof b === 'object') {
    return (a as { error?: string }).error === (b as { error?: string }).error;
  }
  return a === b;
}
