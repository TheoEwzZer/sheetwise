// Excel functions missing from fast-formula-parser, plus a fix for comparisons with empty cells.
import FormulaParser from 'fast-formula-parser';
import operators from 'fast-formula-parser/formulas/operators.js';

interface Arg {
  value: unknown;
  isArray?: boolean;
  isRangeRef?: boolean;
  omitted?: boolean;
}

interface ParsedCriteria {
  op: string;
  value: unknown;
  match?: boolean;
}

const { Criteria } = FormulaParser as unknown as { Criteria: { parse(c: unknown): ParsedCriteria } };
const Infix = (operators as { Infix: { compareOp: (...a: unknown[]) => boolean } }).Infix;

// Excel: an empty cell equals "" (and 0). fast-formula-parser always converts it to 0.
const originalCompare = Infix.compareOp;
let patched = false;
export function patchEmptyComparison(): void {
  if (patched) return;
  patched = true;
  Infix.compareOp = (v1: unknown, op: unknown, v2: unknown, a1: unknown, a2: unknown) => {
    if (v1 == null && typeof v2 === 'string') v1 = '';
    if (v2 == null && typeof v1 === 'string') v2 = '';
    return originalCompare(v1, op, v2, a1, a2);
  };
}

const err = (code: string) => new FormulaParser.FormulaError(code);

function grid(a: Arg): unknown[][] {
  const v = a.value;
  if (Array.isArray(v)) return Array.isArray(v[0]) ? (v as unknown[][]) : [v as unknown[]];
  return [[v]];
}

function flat(a: Arg): unknown[] {
  return grid(a).flat();
}

function scalar(a: Arg | undefined): unknown {
  if (!a) return undefined;
  const v = a.value;
  if (Array.isArray(v)) return Array.isArray(v[0]) ? v[0][0] : v[0];
  return v;
}

function matcher(c: Arg): (v: unknown) => boolean {
  const parsed = Criteria.parse(scalar(c));
  return (v: unknown) => {
    if (parsed.op === 'wc') return parsed.match === (parsed.value as RegExp).test(String(v ?? ''));
    try {
      return Infix.compareOp(v, parsed.op, parsed.value, false, false);
    } catch {
      return false;
    }
  };
}

/** (row, column) indices satisfying every range/criteria pair. */
function matchingCells(pairs: Arg[]): [number, number][] | Error {
  if (pairs.length < 2 || pairs.length % 2) return err('#VALUE!') as unknown as Error;
  const ranges: unknown[][][] = [];
  const tests: ((v: unknown) => boolean)[] = [];
  for (let i = 0; i < pairs.length; i += 2) {
    ranges.push(grid(pairs[i]));
    tests.push(matcher(pairs[i + 1]));
  }
  const h = ranges[0].length;
  const w = ranges[0][0]?.length ?? 0;
  if (ranges.some((r) => r.length !== h || (r[0]?.length ?? 0) !== w)) return err('#VALUE!') as unknown as Error;
  const out: [number, number][] = [];
  for (let r = 0; r < h; r++) {
    for (let c = 0; c < w; c++) {
      if (tests.every((t, k) => t(ranges[k][r][c]))) out.push([r, c]);
    }
  }
  return out;
}

function aggregateIfs(target: Arg, pairs: Arg[], agg: (nums: number[]) => unknown): unknown {
  const cells = matchingCells(pairs);
  if (!Array.isArray(cells)) return cells;
  const g = grid(target);
  const nums: number[] = [];
  for (const [r, c] of cells) {
    const v = g[r]?.[c];
    if (typeof v === 'number') nums.push(v);
  }
  return agg(nums);
}

function same(a: unknown, b: unknown): boolean {
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return a === b;
}

function toText(v: unknown): string {
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

export const EXTRA_FUNCTIONS: Record<string, (...args: Arg[]) => unknown> = {
  COUNTIFS: (...pairs) => {
    const cells = matchingCells(pairs);
    return Array.isArray(cells) ? cells.length : cells;
  },
  SUMIFS: (sum, ...pairs) => aggregateIfs(sum, pairs, (n) => n.reduce((a, b) => a + b, 0)),
  AVERAGEIFS: (avg, ...pairs) =>
    aggregateIfs(avg, pairs, (n) => (n.length ? n.reduce((a, b) => a + b, 0) / n.length : err('#DIV/0!'))),
  MAXIFS: (max, ...pairs) => aggregateIfs(max, pairs, (n) => (n.length ? Math.max(...n) : 0)),
  MINIFS: (min, ...pairs) => aggregateIfs(min, pairs, (n) => (n.length ? Math.min(...n) : 0)),
  TEXTJOIN: (delim, ignoreEmpty, ...items) => {
    const d = toText(scalar(delim));
    const skip = !!scalar(ignoreEmpty);
    const parts: string[] = [];
    for (const it of items) {
      for (const v of flat(it)) {
        const t = toText(v);
        if (skip && t === '') continue;
        parts.push(t);
      }
    }
    return parts.join(d);
  },
  SWITCH: (expr, ...rest) => {
    const v = scalar(expr);
    let i = 0;
    for (; i + 1 < rest.length; i += 2) {
      if (same(v, scalar(rest[i]))) return scalar(rest[i + 1]);
    }
    return i < rest.length ? scalar(rest[i]) : err('#N/A');
  },
  VALUE: (text) => {
    const v = scalar(text);
    if (typeof v === 'number') return v;
    const s = toText(v).trim().replace(/\s/g, '').replace(',', '.');
    const pct = s.endsWith('%');
    const n = Number(pct ? s.slice(0, -1) : s);
    if (!s || !Number.isFinite(n)) return err('#VALUE!');
    return pct ? n / 100 : n;
  },
  XLOOKUP: (lookup, lookupArray, returnArray, ifNotFound, matchMode) => {
    const needle = scalar(lookup);
    const mode = Number(scalar(matchMode) ?? 0);
    const hay = flat(lookupArray);
    const g = grid(returnArray);
    let idx = -1;
    if (mode === 0 || mode === 2) {
      idx = hay.findIndex((v) => same(v, needle));
    } else {
      // -1: exact match or next smaller; 1: exact match or next larger
      let best: number | undefined;
      hay.forEach((v, i) => {
        if (typeof v !== typeof needle) return;
        if (same(v, needle)) {
          if (idx < 0 || best !== undefined) idx = i;
          best = undefined;
          return;
        }
        const ok = mode < 0 ? (v as number) < (needle as number) : (v as number) > (needle as number);
        if (!ok || (idx >= 0 && best === undefined)) return;
        if (best === undefined || (mode < 0 ? (v as number) > best : (v as number) < best)) {
          best = v as number;
          idx = i;
        }
      });
    }
    if (idx < 0) return ifNotFound && !ifNotFound.omitted ? scalar(ifNotFound) : err('#N/A');
    const vertical = grid(lookupArray).length > 1;
    if (vertical) return g[idx]?.[0] ?? err('#N/A');
    return g[0]?.[idx] ?? err('#N/A');
  },
};
